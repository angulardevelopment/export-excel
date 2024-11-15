import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-excel-document-response',
  templateUrl: './excel-document-response.component.html',
  styleUrls: ['./excel-document-response.component.scss']
})
export class ExcelDocumentResponseComponent implements OnInit {

  constructor(public http: HttpClient) { }

  ngOnInit(): void {
  }

  svdfd(){
    this.http.post<any>(apiEndpoint,request,{ responseType: 'blob'} )
  }

  downloadFile(data: Response) {
    const fileName= 'test';
    const blob = new Blob([data], { type: 'application/octet-stream' });
    saveAs(blob, fileName + '.xlsx');
    // const url= window.URL.createObjectURL(blob);
    // window.open(url);


  }
}
