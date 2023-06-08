import { useState, useEffect } from "react";
import { read, utils, writeFile } from "xlsx";
import { Box, Button, ButtonGroup, Typography } from "@mui/material";

interface XlsxParserT {
  isCut: string;
  typeMaterial: string;
  material: string;
  codeMaterial: string;
  depth: number;
  position: string;
  name: string;
  finishedLength: number;
  finishedWidth: number;
  cutLength: number;
  cutWidth: number;
  qty: number;
  isOrientation: string;
  groove: string;
  l1Name: string;
  l1Designation: string;
  l1Depth: number;
  l2Name: string;
  l2Designation: string;
  l2Depth: number;
  w1Name: string;
  w1Designation: string;
  w1Depth: number;
  w2Name: string;
  w2Designation: string;
  w2Depth: number;
  priority: string;
  comment: string;
  originalMaterial: string;
  claddingPlastyTop1: string;
  claddingPlastyBottom1: string;
}

export const CuttingAssistent = () => {
  const [xlsxParser, setXlsxParser] = useState<XlsxParserT[]>([]);
  const [fileNameDisplay, setFileNameDisplay] = useState("");
  const [fileName, setFileName] = useState("");

  const [cuttingsString, setCuttingsString] = useState<XlsxParserT[]>([]);
  const [detailsBig, setDetailsBig] = useState<XlsxParserT[]>([]);
  const [detailsMini, setDetailsMini] = useState<XlsxParserT[]>([]);

  useEffect(() => {
    return () => {
      setXlsxParser([]);
      setFileName("");
    };
  }, []);

  useEffect(() => {
    setCuttingsString(xlsxParser);
  }, [xlsxParser]);

  const column: string[] = [
    "isCut",
    "typeMaterial",
    "material",
    "codeMaterial",
    "depth",
    "position",
    "name",
    "finishedLength",
    "finishedWidth",
    "cutLength",
    "cutWidth",
    "qty",
    "isOrientation",
    "groove",
    "l1Name",
    "l1Designation",
    "l1Depth",
    "l2Name",
    "l2Designation",
    "l2Depth",
    "w1Name",
    "w1Designation",
    "w1Depth",
    "w2Name",
    "w2Designation",
    "w2Depth",
    "priority",
    "comment",
    "originalMaterial",
    "claddingPlastyTop1",
    "claddingPlastyBottom1",
  ];

  const handleFileAsync = async (e: any) => {
    const file = e.target.files[0];
    setFileName(file.name.split(".")[0]);
    setFileNameDisplay(file.name);
    const checkExtension = file.name.split(".").slice(-1)[0];
    if (checkExtension === "xlsx" || checkExtension === "xls") {
      const data = await file.arrayBuffer();
      const workbook = read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const articlesAddName = utils.sheet_add_aoa(worksheet, [column], {
        origin: "A1",
      });
      const articlesData: any = utils.sheet_to_json(articlesAddName);
      setXlsxParser(articlesData);
    } else {
      console.error("Неверный формат файла");
    }
  };

  const sortCutStrings = (cuttingsString: XlsxParserT[]) => {
    const detailsSortBig: XlsxParserT[] = [];
    const detailsSortMini: XlsxParserT[] = [];

    cuttingsString.forEach((string) => {
      if (
        string.finishedLength < 140 &&
        (string.l1Designation !== "" || string.l2Designation !== "")
      ) {
        detailsSortMini.push(string);
      } else if (
        string.finishedWidth < 140 &&
        (string.w1Designation !== "" || string.w2Designation !== "")
      ) {
        detailsSortMini.push(string);
      } else {
        detailsSortBig.push(string);
      }
    });
    console.log("detailsSortBig", detailsSortBig);
    console.log("detailsSortMini", detailsSortMini);
    setDetailsBig(detailsSortBig);
    setDetailsMini(detailsSortMini);
  };

  const downloadSortCutting = (rows: XlsxParserT[], name: string): void => {
    const rowsCalc: XlsxParserT[] = [];
    rows.forEach((el: XlsxParserT) => rowsCalc.push(el));

    const worksheet = utils.json_to_sheet(rowsCalc);
    const workbook = utils.book_new();
    utils.book_append_sheet(workbook, worksheet);
    utils.sheet_add_aoa(
      worksheet,
      [
        [
          "Кроить",
          "Тип материала",
          "Материал",
          "Артикул материала",
          "Толщина",
          "Позиция",
          "Наименования",
          "Длина готовая",
          "Ширина готовая",
          "Длина распиловочная",
          "Ширина распиловочная",
          "Кол-во",
          "Ориентация",
          "Паз",
          "L1 - Наим.",
          "L1 - Обозн.",
          "L1 - Толщина",
          "L2 - Наим.",
          "L2 - Обозн.",
          "L2 - Толщина",
          "W1 - Наим.",
          "W1 - Обозн.",
          "W1 - Толщина",
          "W2 - Наим.",
          "W2 - Обозн.",
          "W2 - Толщина",
          "Приоритет",
          "Комментарий",
          "%Оригинальный материал",
          "%Облицовка пласти верх 1",
          "%Облицовка пласти низ 1",
        ],
      ],
      {
        origin: "A1",
      }
    );

    const max_width = rowsCalc.reduce((w, r) => Math.max(w, r.name.length), 10);
    worksheet["!cols"] = [
      { wch: max_width },
      { wch: 8 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
    ];

    writeFile(workbook, `${fileName.split("0")[0]}_${name}.xlsx`, {
      compression: true,
    });
  };

  return (
    <Box>
      <Box>
        <ButtonGroup>
          <Button variant="contained" size={"small"} component="label">
            Загрузить файл .XLSX
            <input type="file" onChange={(e) => handleFileAsync(e)} hidden />
          </Button>
          <Button
            disabled={!xlsxParser.length}
            onClick={() => {
              console.log("Log xlsxParserFile", xlsxParser);
            }}
            variant="contained"
            size={"small"}
            component="label"
          >
            log
          </Button>
        </ButtonGroup>

        <Box style={{ fontSize: "13px" }}>
          {xlsxParser.length ? (
            <Typography>Файл: {fileNameDisplay}</Typography>
          ) : null}
        </Box>
      </Box>
      <Box>
        <Button
          sx={{ mt: 1 }}
          size={"small"}
          variant={"contained"}
          onClick={() => console.log("cuttingsString", cuttingsString)}
        >
          Log cuttingsString
        </Button>
        <Button
          sx={{ mt: 1, ml: 1 }}
          size={"small"}
          variant={"contained"}
          onClick={() => sortCutStrings(cuttingsString)}
          disabled={!cuttingsString.length}
        >
          sort Cut Strings
        </Button>
        <Button
          sx={{ mt: 1, ml: 1 }}
          size={"small"}
          variant={"contained"}
          onClick={() => {
            downloadSortCutting(detailsMini, "на полосы");
            downloadSortCutting(detailsBig, "большие детали");
          }}
          disabled={!detailsMini.length}
        >
          Cut DownLoad
        </Button>
      </Box>
    </Box>
  );
};
