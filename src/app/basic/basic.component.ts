import { Component, OnInit } from '@angular/core';
import { saveAs } from "file-saver";
import { utils, write, WorkBook } from "xlsx";
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import { GoogleSheetsDbService } from 'ng-google-sheets-db';
import { Observable } from 'rxjs';
import { HttpClient } from '@angular/common/http';
type AOA = Array<Array<any>>; // type AOA = any[][];
@Component({
  selector: 'app-basic',
  templateUrl: './basic.component.html',
  styleUrls: ['./basic.component.scss']
})
export class BasicComponent implements OnInit {
  ws_data = [
    ["S", "h", "e", "e", "t", "J", "S"],
    [1, 2, 3, 4, 5]
  ];
  data2 = [{
    eid: 'e101',
    ename: 'ravi',
    esal: 1000
  },
  {
    eid: 'e102',
    ename: 'ram',
    esal: 2000
  },
  {
    eid: 'e103',
    ename: 'rajesh',
    esal: 3000
  }];
  data: AOA = [[1, 2], [3, 4]];

  characters$: Observable<any[]>;
  willDownload = false;

  constructor(private googleSheetsDbService: GoogleSheetsDbService, public http: HttpClient) { }

  ngOnInit(): void {

}

googleSheetUsage(){
  const characterAttributesMapping = {
    id: 'ID',
    name: 'Name',
    name2: 'data',

    email: 'Email Address',
    contact: {
      _prefix: 'Contact',
      street: 'Street',
      streetNumber: 'Street Number',
      zip: 'ZIP',
      city: 'City'
    },
    skills: {
      _prefix: 'Skill',
      _listField: true
    }
  };

  // not use this publish url
  // https://docs.google.com/spreadsheets/d/e/2PACX-1vTVJ2OLVjj_3O31RZaTg4nFcBMBqBUEWXGBo81uJ7IdcmKIo3PGIPNB9-yWzGnaQjr0nkN6FXLo_b1q/pubhtml

 // use th edit URL
  // https://docs.google.com/spreadsheets/d/16j0t7K1EwBRTHERly4Ggp-ZVAgwlXYCnJDAYFGJ1vHk/edit
  this.characters$ = this.googleSheetsDbService.get('16j0t7K1EwBRTHERly4Ggp-ZVAgwlXYCnJDAYFGJ1vHk', 1, characterAttributesMapping);

// this.characters$ = this.googleSheetsDbService.getActive(
//   '1gSc_7WCmt-HuSLX01-Ev58VsiFuhbpYVo8krbPCvvqA', 1, characterAttributesMapping, 'Active');

this.characters$.subscribe((res)=>{
console.log(res, 'data');

})
}

// public getSheetData(): Observable<any> {
//   const sheetId = '15Kndr-OcyCUAkBUcq6X3BMqKa_y2fMAXfPFLiSACiys';
//   const url = `https://spreadsheets.google.com/feeds/list/${sheetId}/od6/public/values?alt=json`;

//   return this.http.get(url)
//     .pipe(map((res: any) => {
//         const data = res.feed.entry;

//         const returnArray: Array<any> = [];
//         if (data && data.length > 0) {
//           data.forEach(entry => {
//             const obj = {};
//             for (const x in entry) {
//               if (x.includes('gsx$') && entry[x].$t) {
//                 obj[x.split('$')[1]] = entry[x]['$t'];
//               }
//             }
//             returnArray.push(obj);
//           });
//         }
//         return returnArray;
//       })
//     );
// }
  // generate workbook and add the worksheet
  exportSheet() {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.ws_data);
    // other services
    // XLSX.utils.table_to_sheet
    // > console.log(XLSX.utils.sheet_to_csv(ws));
    // > console.log(XLSX.utils.sheet_to_csv(ws, {FS:"\t"})); S    h    e    e    t    J    S 1    2    3    4    5    6    7 2    3    4    5    6    7    8
    // > console.log(XLSX.utils.sheet_to_csv(ws,{FS:":",RS:"|"})); S:h:e:e:t:J:S|1:2:3:4:5:6:7|2:3:4:5:6:7:8|
    // var o = XLSX.utils.sheet_to_formulae(ws);
    // o.filter(function (v, i) { return i % 5 === 0; });
    // ['A1=\'S', 'F1=\'J', 'D2=4', 'B3=3', 'G3=8']
    // new empty workbook
    const wb: XLSX.WorkBook = XLSX.utils.book_new();// create a new blank book
    // to add data and sheet
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');// add the worksheet to the book
    // Another way to generate excel
    // var wb = XLSX.utils.table_to_book(document.getElementById('tbl')); // worksheet or workbook

    // add little Comment
    // if (!ws.A1.c) ws.A1.c = []
    // ws.A1.c.push({ a: "SheetJS", t: "I'm a little comment, short and stout!" });
    // Formatted text
    ws['B1'].w = "3";
    //To add Hyperlink
    ws['A2'].l = { Target: "#", Tooltip: "Find us @ SheetJS.com!" };
    // to download file first way
    // XLSX.writeFile(wb, 'epltable.xlsx'); //  attempts to write  wb  to  filename // initiate a file download in browser
    // var workbook = XLSX.readFile('test.xlsx');

    // save to file    // to download file second way
    const wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'binary' };
    const wbout: string = XLSX.write(wb, wopts); // attempts to write the workbook  wb, write_opts- { bookType: 'xlsx', type: 'binary' } // type- array, 'binary'
    //    the saveAs call downloads a file on the local machine
    saveAs(new Blob([this.s2ab(wbout)]), 'test.xlsx');//  saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx");

  }



  // To preview the excel data
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      /* grab first sheet */
      const wsname: string = wb.SheetNames[0]; // first_sheet_name
      const ws: XLSX.WorkSheet = wb.Sheets[wsname]; // worksheet
      //   var address_of_cell = 'A1';
      //   var desired_cell = ws[address_of_cell];
      //   var desired_value = desired_cell.v;
      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
      console.log(this.data);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  onFileChange2(ev) {
    let workBook = null;
    let jsonData = null;
    const reader = new FileReader();
    const file = ev.target.files[0];
    reader.onload = (event) => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary' });
      jsonData = workBook.SheetNames.reduce((initial, name) => {
        const sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(sheet);
        return initial;
      }, {});
      const dataString = JSON.stringify(jsonData);
      console.log(dataString, jsonData,jsonData.Sheet1, 'jsonData');

      document.getElementById('output').innerHTML = dataString.slice(0, 300).concat("...");
      this.setDownload(dataString);
    }
    reader.readAsBinaryString(file);
  }


  setDownload(data) {
    this.willDownload = true;
    setTimeout(() => {
      const el = document.querySelector("#download");
      el.setAttribute("href", `data:text/json;charset=utf-8,${encodeURIComponent(data)}`);
      el.setAttribute("download", 'xlsxtojson.json');
    }, 1000)
  }


  public exportAsExcelFile(): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.data2); // generate worksheet
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    // type can be buffer also. If we are using xlsx-style for styling then we can also replace
    // XLSX with XLSXStyle.
    const data: Blob = new Blob([excelBuffer], {
      type: ''
    });
    FileSaver.saveAs(data, 'excelFileName' + '_export_' + new Date().getTime() + '.xlsx');
  }
  generateExcelsheet() {
    const ws_name = "SomeSheet";
    const wb: WorkBook = { SheetNames: [], Sheets: {} };
    const ws: any = utils.json_to_sheet(this.data2, {
      header: ["First", "Second", "A"],
      skipHeader: true
    });
    // Add the sheet name to the list
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    const wbout = write(wb, { bookType: "xls", bookSST: true, type: "binary" });


    saveAs(
      new Blob([this.s2ab(wbout)], { type: "application/octet-stream" }),
      "exported.xls"
    );
  }
  s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
  }

  // To get excel data from assets you can use below approach-
  fetchDataFromAssets() {
    const testUrl = "../assets/demo.xlsx";


    var oReq = new XMLHttpRequest();
    oReq.open("GET", testUrl, true);
    oReq.responseType = "arraybuffer";
    oReq.onload = function (e) {
      var arraybuffer = oReq.response;
      /* convert data to binary string */
      var data = new Uint8Array(arraybuffer);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i) {
        arr[i] = String.fromCharCode(data[i]);
      }
      var bstr = arr.join("");
      //        Call XLSX
      var workbook = XLSX.read(bstr, { type: "binary" });
      //  DO SOMETHING WITH workbook HERE
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      var json = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
        { header: 1, raw: true }
      );
      var jsonOut = JSON.stringify(json);
      console.log("test", json);
    };
    oReq.send();
  }

  fetchCSVData() {
    const testUrl = "/assets/cm01JUL2021bhav.csv";

    var oReq = new XMLHttpRequest();
    oReq.open("GET", testUrl, true);
    oReq.responseType = "arraybuffer";
    oReq.onload = function (e) {
      var arraybuffer = oReq.response;
      /* convert data to binary string */
      var data = new Uint8Array(arraybuffer);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i) {
        arr[i] = String.fromCharCode(data[i]);
      }
      var bstr = arr.join("");
      //        Call XLSX
      var workbook = XLSX.read(bstr, { type: "binary" });
      //  DO SOMETHING WITH workbook HERE
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      var json = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
        { header: 1, raw: true }
      );
      var jsonOut = JSON.stringify(json);
      console.log("test", json, jsonOut);
    };
    oReq.send();
  }

  // If you are using xlsxstyle then you can use these methods
  // private wrapAndCenterCell(cell: XLSX.CellObject) {
  //   const wrapAndCenterCellStyle = {
  //     alignment: { wrapText: true, vertical: "center", horizontal: "center" }
  //   };
  //   this.setCellStyle(cell, wrapAndCenterCellStyle);
  // }
  // private setCellStyle(cell: XLSX.CellObject, style: {}) {
  //   cell.s = style;
  // }
  //   setFormulaOnCell (worksheet) {
  //   worksheet['C1'] = { t:'n', f: "SUM(A1:A3*B1:B3)", F:"C1:C1" };  //setting formula
  //  }
  //   createWorksheet(wb) {



  //      var i = 0;
  //             for(i = 0; i !== this.data.length; ++i) console.log(this.data[i]);

  //      if(!wb.Props) wb.Props = {};
  //     // // wb.Props.Title = "style";
  //     if(!wb.Custprops)
  //     wb.Custprops={};
  //     wb.Custprops["Custom Property"] ="Custom Value";
  //     XLSX.write(wb,{Props:{Author:"SheetJS"}});

  //   }
  // set A2 formatted string value
  //styling properties
  //        vertAlign: "superscript",
  //     "numFmt": "0%"
  //  "numFmt": "0.00%;\\(0.00%\\);\\-;@",
  //   "numFmt": "0.0000%"
  //     "numFmt": "d-mmm-yy"
  // horizontal-left,right,center
  // vertical-top,bottom,center
  // alignment": {
  //               "indent": "1"
  // "numFmt": "d-mmm-yy"
  //  "numFmt": "h:mm:ss AM/PM"
  //           },
  //      "numFmt": "m/d/yy"
  //      "numFmt": "m/d/yy"
  //   "numFmt": "m/d/yy h:mm:ss AM/PM"

}
