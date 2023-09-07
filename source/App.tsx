import {
  FieldType,
  IOpenSegment,
  IOpenSegmentType,
  IOpenTextSegment,
  IWidgetTable,
  IWidgetView,
  ViewType,
  bitable,
  IAddFieldConfig,
} from "@base-open/web-api";
import { Button, Popconfirm, Select, Upload } from "@douyinfe/semi-ui";
import "./App.css";
import { useCallback, useEffect, useState } from "react";
import { Spin, Toast, Table } from "@douyinfe/semi-ui";
import * as XLSX from "xlsx";
import Column from "@douyinfe/semi-ui/lib/es/table/Column";
import Text from "@douyinfe/semi-ui/lib/es/typography/text";

async function getActiveTableAndView() {
  const selection = await bitable.base.getSelection();
  if (!selection.tableId) {
    return null;
  }
  const table: IWidgetTable = await bitable.base.getTableById(
    selection.tableId
  );
  if (!selection.viewId) {
    return null;
  }
  const view = await table.getViewById(selection.viewId);
  const type = await view.getType();
  if (type !== ViewType.Grid) {
    return null;
  }
  return { table, view };
}

export interface SheetInfo {
  name: string;
  tableData: {
    fields: {
      name: string;
    }[];
    records: {
      [key: string]: string;
    }[];
  };
}

export interface ExcelDataInfo {
  sheets: SheetInfo[];
}

async function readExcel(fileList: File[]): Promise<ExcelDataInfo | null> {
  const res: ExcelDataInfo = {} as ExcelDataInfo;
  const files = fileList;
  if (!files) {
    return null;
  }
  if (files.length <= 0) {
    return null;
  } else if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
    Toast.warning("上传格式不正确，请上传 xls 或者 xlsx 格式");
    return null;
  }

  const fileReader = new FileReader();
  let resolve: any = () => { };
  fileReader.onload = (ev) => {
    try {
      const result = ev.target?.result;
      const workbook = XLSX.read(result, {
        type: "binary",
      });
      res.sheets = [];
      // need to change here
      for (let name of workbook.SheetNames) {
        const ws = XLSX.utils.sheet_to_json<{ [key: string]: string }>(
          workbook.Sheets[name], { raw: false }
        );

        const fields = [];
        for (let key in ws[0]) {
          fields.push({
            name: key,
          });
        }

        res.sheets.push({
          name,
          tableData: {
            fields,
            records: ws,
          },
        });
      }
      resolve(res);
    } catch (e) {
      resolve(null);
    }
  };
  fileReader.readAsBinaryString(files[0]);

  return new Promise((r) => (resolve = r));
}

async function importTable(
  table: IWidgetTable,
  view: IWidgetView,
  data: SheetInfo
) {
  const { fields, records } = data.tableData; // 即将导入的数据
  const fieldIdMap: { [key: string]: string } = {};
  const fieldMetaList = await view.getFieldMetaList();

  const firstField = fieldMetaList.shift();
  if (!firstField) {
    Toast.error("导入失败，数据表异常");
    return;
  }

  const firstFieldType = firstField.type;
  if (firstFieldType !== FieldType.Text) {
    Toast.error("导入失败，请将第一列字段类型修改为多行文本后再进行导入");
    return;
  }

  // 删除除了第一个字段之外其它所有字段
  await Promise.all([
    fieldMetaList.map((field) => {
      return table.deleteField(field.id);
    }),
  ]);

  // 清空所有记录
  const recordIdList = await table.getRecordIdList();
  await Promise.all(
    recordIdList.map((id) => {
      if (id) {
        return table.deleteRecord(id);
      }
    })
  );

  // 更新第一个字段的名
  await table.setField(firstField.id, {
    name: fields[0].name,
    property: null,
  });
  fieldIdMap[fields[0].name] = firstField.id;

  // 创建新字段
  //function to return field type based on excel field type?
  /** COLUMNS NEEDED
Date: FieldType.DateTime
Name: FieldType.Text
Staff: FieldType.User - with direct manager, etc.
Clock in & out: FieldType.Text
Calculated hours: FieldType.Number
The rest: FieldType.Text
**/

  //function currently not used: returns field type required on base for column in Excel
  function getFieldType(fieldName: string): FieldType {
    let f = fieldName.toLowerCase()
    if (f == "date") {
      return FieldType.Text
    } else if (f == "person") {
      return FieldType.User
    } else if (f == "calculated hours") {
      return FieldType.Number
    }
    else return FieldType.Text;
  }

  for (let i = 1; i < fields.length; i++) {
    var n = fields[i].name;

    const id = await table.addField({
      name: n,
      type: FieldType.Text, //getFieldType(n)
      property: null,
    });
    fieldIdMap[n] = id;
  }

  // 添加记录
  await Promise.all([
    records.map((record) => {
      const v: { [key: string]: IOpenTextSegment[] } = {};
      for (let key in record) {
        v[fieldIdMap[key]] = [
          { type: IOpenSegmentType.Text, text: String(record[key]) },
        ];
      }
      return table.addRecord({
        fields: v,
      });
    }),
  ]);

  Toast.success("导入成功");
}

export function App() {
  const [loading, setLoading] = useState<boolean>(false);
  const [importing, setImporting] = useState<boolean>(false);
  const [activeTableInfo, setActiveTable] = useState<{
    table: IWidgetTable;
    view: IWidgetView;
    tableName: string;
    viewName: string;
  } | null>(null);

  const updateActiveTable = useCallback(() => {
    setLoading(false);
    getActiveTableAndView().then(async (data) => {
      if (data) {
        const [tableName, viewName] = await Promise.all([
          data.table.getName(),
          data.view.getName(),
        ]);
        setLoading(true);
        setActiveTable({
          ...data,
          tableName,
          viewName,
        });
      } else {
        setLoading(true);
        setActiveTable(null);
      }
    });
  }, []);

  useEffect(() => {
    updateActiveTable();
    const unsub = bitable.base.onSelectionChange(() => {
      updateActiveTable();
    });
    return () => {
      unsub();
    };
  }, []);

  const [selected, setSelected] = useState<boolean>(false);
  const [dataInfo, setDataInfo] = useState<ExcelDataInfo | null>(null);
  const [index, setIndex] = useState<number>(0);

  const sheetCount = dataInfo?.sheets.length;

  const selectVisible = sheetCount && sheetCount > 1;

  const selectFile = useCallback((files: File[]) => {
    readExcel(files).then((data) => {
      if (!data) {
        Toast.warning("解析文件失败");
      }
      setDataInfo(data);
      setSelected(true);
    });
  }, []);

  if (!loading) {
    return (
      <div className="loading">
        <Spin size={"large"}></Spin>
      </div>
    );
  }

  if (!activeTableInfo) {
    return (
      <div className="error">
        <span>请打开一个数据表的表格视图</span>
      </div>
    );
  }

  if (importing) {
    return (
      <div className="importing">
        <Spin size={"large"}></Spin>
        <span>导入中</span>
      </div>
    );
  }

  return (
    <div
      style={{
        height: "100%",
      }}
    >
      {!selected && (
        <div className="selectFile">
          <Upload
            draggable={true}
            accept=".xls,.xlsx"
            dragMainText={"点击上传文件或拖拽文件到这里"}
            dragSubText="支持 xls、xlsx 类型文件"
            onFileChange={selectFile}
          ></Upload>
        </div>
      )}
      {selected && dataInfo && (
        <div className="main">
          <div className="previewText">表格数据预览</div>
          {selectVisible && (
            <div>
              <Text className="selectTableText">选择要导入的表</Text>
              <Select
                defaultValue={dataInfo.sheets[0].name}
                style={{ width: 120 }}
                onSelect={(v) => {
                  setIndex(v as any);
                }}
                size="small"
              >
                {dataInfo.sheets.map((sheet, index) => {
                  return (
                    <Select.Option value={index}>{sheet.name}</Select.Option>
                  );
                })}
              </Select>
            </div>
          )}

          <div className="table">
            {dataInfo !== null && <TableView data={dataInfo.sheets[index]} />}
          </div>

          <div className="footer">
            <Button
              onClick={() => {
                setSelected(false);
                setDataInfo(null);
              }}
            >
              重新选择文件
            </Button>
            <Popconfirm
              title="确定要导入吗？"
              content={
                <div>
                  导入后将覆盖{" "}
                  <span style={{ fontWeight: 700, color: "#000" }}>
                    {activeTableInfo.tableName}
                  </span>{" "}
                  中数据
                </div>
              }
              onConfirm={() => {
                setImporting(true);
                importTable(
                  activeTableInfo.table,
                  activeTableInfo.view,
                  dataInfo.sheets[index]
                ).then(() => {
                  setImporting(false);
                });
              }}
            >
              <Button>导入</Button>
            </Popconfirm>
          </div>
        </div>
      )}
    </div>
  );
}

function TableView(props: { data: SheetInfo }) {
  const { data } = props;
  return (
    <div className="tableViewContainer">
      <Table dataSource={data.tableData.records} pagination={false}>
        {data.tableData.fields.map((field) => {
          return (
            <Column
              title={field.name}
              dataIndex={field.name}
              key={field.name}
            />
          );
        })}
      </Table>
    </div>
  );
}
