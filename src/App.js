import React, { Component } from "react";
import DataGrid, {
  Export,
  GroupPanel,
  Column,
} from "devextreme-react/data-grid";
import Button from "devextreme-react/button";
import ExcelJS from "exceljs/dist/es5/exceljs.browser";
import saveAs from "file-saver";
import service from "./data.js";

class App extends Component {
  constructor(props) {
    super(props);
    this.employees = service.getEmployees();

    this.state = {
      excelFilterEnabled: true,
    };
  }

  handleExportClick = (e) => {
    this.excelExport(this.state.instanceDataGrid);
  };

  excelExport = (DataGrid) => {
    var ExcelJSWorkbook = new ExcelJS.Workbook();
    var worksheet = ExcelJSWorkbook.addWorksheet("ExcelJS sheet", {
      pageSetup: {
        horizontalCentered: true,
        verticalCentered: true,
        paperSize: 9,
        orientation: "portrait",
        margins: {
          left: 0.3149606,
          right: 0.3149606,
          top: 0.3543307,
          bottom: 0.3543307,
          header: 0.3149606,
          footer: 0.3149606,
        },
      },
    });

    var columns = DataGrid.getVisibleColumns();

    // var sheet = ExcelJSWorkbook.addWorksheet(sheetName, {
    //   pageSetup: { paperSize: 9, orientation: "landscape" }
    // });

    worksheet.mergeCells("A2:AH2");
    worksheet.getCell("A2").alignment = { horizontal: "center" };

    const customCell = worksheet.getCell("A2");
    worksheet.getRow(2).height = 25;
    worksheet.getRow("A2").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };

    worksheet.mergeCells("A3:AH3");
    worksheet.getCell("A3").alignment = { horizontal: "center" };
    worksheet.getRow(3).height = 25;

    worksheet.mergeCells("B6:B7");
    worksheet.getCell("B6:B7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    worksheet.mergeCells("C6:Z6");
    worksheet.getCell("C6:Z6").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCell1 = worksheet.getCell("A3");
    const customCell2 = worksheet.getCell("A5");
    worksheet.mergeCells("A5:D5");

    worksheet.mergeCells("A6:A7");
    worksheet.getCell("A6:A7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    const customCell3 = worksheet.getCell("A6");
    customCell3.value = "No";

    // worksheet.getCell("B6").width = 135;
    const customCell4 = worksheet.getCell("B6");
    customCell4.value = "Stasiun";

    const customCell5 = worksheet.getCell("C6");
    customCell5.value = "Tanggal";

    const customCellTanggal1 = worksheet.getCell("C7");
    customCellTanggal1.value = "1";
    worksheet.getCell("C7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal2 = worksheet.getCell("D7");
    customCellTanggal2.value = "2";
    worksheet.getCell("D7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal3 = worksheet.getCell("E7");
    customCellTanggal3.value = "3";
    worksheet.getCell("E7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal4 = worksheet.getCell("F7");
    customCellTanggal4.value = "4";
    worksheet.getCell("F7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal5 = worksheet.getCell("G7");
    customCellTanggal5.value = "5";
    worksheet.getCell("G7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal6 = worksheet.getCell("H7");
    customCellTanggal6.value = "6";
    worksheet.getCell("H7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal7 = worksheet.getCell("I7");
    customCellTanggal7.value = "7";
    worksheet.getCell("I7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal8 = worksheet.getCell("J7");
    customCellTanggal8.value = "8";
    worksheet.getCell("J7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal9 = worksheet.getCell("K7");
    customCellTanggal9.value = "9";
    worksheet.getCell("K7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal10 = worksheet.getCell("L7");
    customCellTanggal10.value = "10";
    worksheet.getCell("L7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal11 = worksheet.getCell("M7");
    customCellTanggal11.value = "11";
    worksheet.getCell("M7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal12 = worksheet.getCell("N7");
    customCellTanggal12.value = "12";
    worksheet.getCell("N7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal13 = worksheet.getCell("O7");
    customCellTanggal13.value = "13";
    worksheet.getCell("O7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal14 = worksheet.getCell("P7");
    customCellTanggal14.value = "14";
    worksheet.getCell("P7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal15 = worksheet.getCell("Q7");
    customCellTanggal15.value = "15";
    worksheet.getCell("Q7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal16 = worksheet.getCell("R7");
    customCellTanggal16.value = "16";
    worksheet.getCell("R7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    const customCellTanggal17 = worksheet.getCell("S7");
    customCellTanggal17.value = "17";
    worksheet.getCell("S7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal18 = worksheet.getCell("T7");
    customCellTanggal18.value = "18";
    worksheet.getCell("T7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal19 = worksheet.getCell("U7");
    customCellTanggal19.value = "19";
    worksheet.getCell("U7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal20 = worksheet.getCell("V7");
    customCellTanggal20.value = "20";
    worksheet.getCell("V7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal21 = worksheet.getCell("W7");
    customCellTanggal21.value = "21";
    worksheet.getCell("W7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal22 = worksheet.getCell("X7");
    customCellTanggal22.value = "22";
    worksheet.getCell("X7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal23 = worksheet.getCell("Y7");
    customCellTanggal23.value = "23";
    worksheet.getCell("Y7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal24 = worksheet.getCell("Z7");
    customCellTanggal24.value = "24";
    worksheet.getCell("Z7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal25 = worksheet.getCell("AA7");
    customCellTanggal25.value = "25";
    worksheet.getCell("AA7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal26 = worksheet.getCell("AB7");
    customCellTanggal26.value = "26";
    worksheet.getCell("AB7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal27 = worksheet.getCell("AC7");
    customCellTanggal27.value = "27";
    worksheet.getCell("AC7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal28 = worksheet.getCell("AD7");
    customCellTanggal28.value = "28";
    worksheet.getCell("AD7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal29 = worksheet.getCell("AE7");
    customCellTanggal29.value = "29";
    worksheet.getCell("AE7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal30 = worksheet.getCell("AF7");
    customCellTanggal30.value = "30";
    worksheet.getCell("AF7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    const customCellTanggal31 = worksheet.getCell("AG7");
    customCellTanggal31.value = "31";
    worksheet.getCell("AG7").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    worksheet.mergeCells("AH6:AH7");
    worksheet.getCell("AH6:AH7").alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    worksheet.columns = [{ key: "AH", width: 250 }];

    const customCellRataDataMasuk = worksheet.getCell("AH6");
    customCellRataDataMasuk.value = "Rata-rata Data Masuk";

    worksheet.mergeCells("AI6:AI7");
    worksheet.getCell("AI6:AI7").alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    const customCellPenyebabDataTidakMasuk = worksheet.getCell("AI6");
    customCellPenyebabDataTidakMasuk.value = "Penyebab Data Tidak Masuk";

    // STYLE CELL HEADER
    customCell.font = {
      name: "Times New Roman",
      family: 4,
      size: 20,
      // underline: true,
      alignment: "center",
      bold: true,
    };

    customCell1.font = {
      name: "Times New Roman",
      family: 4,
      size: 20,
      // underline: true,
      bold: true,
    };
    customCell2.font = {
      name: "Times New Roman",
      family: 4,
      size: 14,
      // underline: true,
      bold: true,
    };
    customCell.value = "Hasil Monitoring BMKGSoft Bulan Mei 2022";
    customCell1.value = "Balai Besar MKG Wilayah 1";
    customCell2.value = "Tipe FORM: Me-48";

    // var headerRow = worksheet.addRow();
    // var headerRow1 = worksheet.addRow();
    // var headerRow2 = worksheet.addRow();

    worksheet.getRow(4).font = { bold: true };

    for (let i = 0; i < columns.length; i++) {
      let currentColumnWidth = DataGrid.option().columns[i].width;
      worksheet.getColumn(i + 1).width =
        currentColumnWidth !== undefined ? currentColumnWidth / 6 : 20;
      // let cell = headerRow.getCell(i + 1);
      // let cell1 = headerRow1.getCell(i + 1);
      // let cell2 = headerRow2.getCell(i + 1);

      // cell.value = columns[i].caption;
      // cell1.value = columns[i].caption;
      // cell2.value = columns[i].caption;
    }

    if (this.state.excelFilterEnabled === true) {
      worksheet.autoFilter = {
        from: {
          row: 3,
          column: 1,
        },
        to: {
          row: 3,
          column: columns.length,
        },
      };
    }

    // eslint-disable-next-line no-unused-expressions
    this.state.excelFilterEnabled === true
      ? (worksheet.views = [{ state: "frozen", ySplit: 7 }])
      : undefined;

    worksheet.properties.outlineProperties = {
      summaryBelow: false,
      summaryRight: false,
    };
    worksheet.properties.outlineProperties = {
      summaryBelow: false,
      summaryRight: false,
    };

    DataGrid.getController("data")
      .loadAll()
      .then(function (allItems) {
        for (let i = 0; i < allItems.length; i++) {
          var dataRow = worksheet.addRow();
          if (allItems[i].rowType === "data") {
            dataRow.outlineLevel = 1;
          }
          for (let j = 0; j < allItems[i].values.length; j++) {
            let cell = dataRow.getCell(j + 1);
            cell.value = allItems[i].values[j];
          }
        }

        const rowCount = worksheet.rowCount;
        // const rowCount1 = worksheet.rowCount;

        // worksheet.mergeCells(`A${rowCount}:I${rowCount + 1}`);
        // worksheet.mergeCells(`B${rowCount}:C${rowCount}`);

        // worksheet.mergeCells(`B${rowCount1}`);

        // worksheet.getRow(1).font = { bold: true };
        worksheet.getCell(`A${rowCount}`).font = {
          name: "Times New Roman",
          family: 4,
          size: 12,
          // bold: true
        };
        // worksheet.getCell(`B${rowCount1}`).font = {
        //   name: "Times New Roman",
        //   family: 4,
        //   size: 12
        //   // underline: true,
        //   // bold: true
        // };
        worksheet.getRow(9).height = 30;
        worksheet.getRow(10).height = 30;
        // worksheet.getCell("C6:Z6").alignment = {
        // worksheet.getRange("B1:E22");

        // worksheet.getRange("B9:B100").alignment = {
        //   horizontal: "left",
        //   vertical: "middle",
        //   wrapText: true
        // };
        worksheet.getCell(`A${rowCount + 1}:AH${rowCount + 1}`).alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
        worksheet.getCell(`A${rowCount + 1}`).font = { bold: true };
        worksheet.mergeCells(`A${rowCount + 1}:AG${rowCount + 1}`);
        worksheet.getCell(`A${rowCount + 1}:AG${rowCount + 1}`).value = "TOTAL";

        // const customCellTanggal25 = worksheet.getCell("Z7");
        // customCellTanggal25.value = "25";
        worksheet.getCell(`A${rowCount}:AH${rowCount}`).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        worksheet.mergeCells(`A${rowCount + 4}:D${rowCount + 4}`);
        worksheet.getCell(`A${rowCount + 4}`).value = "Mengetahui.";
        worksheet.mergeCells(`A${rowCount + 5}:D${rowCount + 5}`);
        worksheet.getCell(`A${rowCount + 5}`).value = "Sub Koordinator";
        worksheet.mergeCells(`A${rowCount + 6}:D${rowCount + 6}`);
        worksheet.getCell(`A${rowCount + 6}`).value =
          "Bidang Manajemen Database MKG,";
        worksheet.mergeCells(`AC${rowCount + 5}:AG${rowCount + 5}`);
        worksheet.getCell(`AC${rowCount + 5}`).value = "Jakarta, ";
        worksheet.mergeCells(`AC${rowCount + 6}:AG${rowCount + 6}`);
        worksheet.getCell(`AC${rowCount + 6}`).value = "Pembuat Laporan, ";

        ExcelJSWorkbook.xlsx.writeBuffer().then(function (buffer) {
          saveAs(
            new Blob([buffer], { type: "application/octet-stream" }),
            `${DataGrid.option().export.fileName}.xlsx`
          );
        });
      });
  };

  render() {
    return (
      <div>
        <DataGrid
          id={"gridContainer"}
          dataSource={this.employees}
          showBorders={true}
          onCellPrepared={this.onCellPrepared}
          onContentReady={this.onContentReady}
          remoteOperations={true}
          // wrapText={true}
          wordWrapEnabled={true}
          // allowColumnReordering={true}
        >
          <Column dataField={"Prefix"} caption={"No"} width={20} />
          <Column dataField={"Stasiun"} width={105} />

          {/* <Column caption="tanggal" width={900}> */}
          <Column dataField={"1"} width={30} />
          <Column dataField={"2"} width={30} />
          <Column dataField={"3"} width={30} />
          <Column dataField={"4"} width={30} />
          <Column dataField={"5"} width={30} />
          <Column dataField={"6"} width={30} />
          <Column dataField={"7"} width={30} />
          <Column dataField={"8"} width={30} />
          <Column dataField={"9"} width={30} />
          <Column dataField={"10"} width={30} />
          <Column dataField={"11"} width={30} />
          <Column dataField={"12"} width={30} />
          <Column dataField={"13"} width={30} />
          <Column dataField={"14"} width={30} />
          <Column dataField={"15"} width={30} />
          <Column dataField={"16"} width={30} />
          <Column dataField={"17"} width={30} />
          <Column dataField={"18"} width={30} />
          <Column dataField={"19"} width={30} />
          <Column dataField={"20"} width={30} />
          <Column dataField={"21"} width={30} />
          <Column dataField={"22"} width={30} />
          <Column dataField={"23"} width={30} />
          <Column dataField={"24"} width={30} />
          <Column dataField={"25"} width={30} />
          <Column dataField={"26"} width={30} />
          <Column dataField={"27"} width={30} />
          <Column dataField={"28"} width={30} />
          <Column dataField={"29"} width={30} />
          <Column dataField={"30"} width={30} />
          <Column dataField={"31"} width={30} />
          {/* </Column> */}

          <Column
            getVisibleColumns={false}
            dataField={"Rata - rata Data Masuk"}
            width={70}
          />
          <Column
            getVisibleColumns={false}
            dataField={"Penyebab Data Tidak Masuk"}
          />

          {/* <Column    dataField={"Position"} width={130} />
          <Column    dataField={"BirthDate"} dateType={"date"} width={130} />
          <Column    dataField={"HireDate"} dateType={"date"} width={100} />
          <Column
            dataField={"SaleAmount"}
            alighment={"right"}
            format={"currency"}
          /> */}
          <Export
            enabled={true}
            fileName="BBW I_Mei 2022_Final"
            excelFilterEnabled={true}

            // customizeExcelCell={this.customizeExcelCell}
          />
          <GroupPanel visible={true} />
        </DataGrid>
        <Button text="export" type="danger" onClick={this.handleExportClick} />
      </div>
    );
  }

  onContentReady = (e) => {
    var instanceGrid = e.component.instance();

    this.setState({
      excelFilterEnabled: instanceGrid.option().export.excelFilterEnabled,
      instanceDataGrid: instanceGrid,
    });
  };

  onCellPrepared(e) {
    if (e.rowType === "data") {
      if (e.data.OrderDate < new Date(2014, 2, 3)) {
        e.cellElement.classList.add("oldOrder");
      }
      if (e.data.SaleAmount > 15000) {
        if (e.column.dataField === "Employee") {
          e.cellElement.classList.add("highAmountOrder_employee");
        }
        if (e.column.dataField === "SaleAmount") {
          e.cellElement.classList.add("highAmountOrder_saleAmount");
        }
      }
    }
  }

  customizeExcelCell(options) {
    if (options.gridCell.rowType === "data") {
      if (options.gridCell.data.SaleAmount > 15000) {
        if (options.gridCell.column.dataField === "Employee") {
          options.font.bold = true;
        }
        if (options.gridCell.column.dataField === "SaleAmount") {
          options.backgroundColor = "#FFBB00";
          options.font.color = "#000000";
        }
      }
    }
  }
}

export default App;
