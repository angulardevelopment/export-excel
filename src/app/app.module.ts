import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { BasicComponent } from './basic/basic.component';
import { ExceljsComponent } from './exceljs/exceljs.component';
import { GoogleSheetsDbService } from 'ng-google-sheets-db';
import { HttpClientModule } from '@angular/common/http';
import { ReactiveFormsModule } from '@angular/forms';
import { ExcelDocumentResponseComponent } from './excel-document-response/excel-document-response.component';

@NgModule({
  declarations: [
    AppComponent,
    BasicComponent,
    ExceljsComponent,
    ExcelDocumentResponseComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    HttpClientModule,
    ReactiveFormsModule
  ],
  providers: [GoogleSheetsDbService ],
  bootstrap: [AppComponent]
})
export class AppModule { }
