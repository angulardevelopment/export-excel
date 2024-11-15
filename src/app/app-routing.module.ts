import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { BasicComponent } from './basic/basic.component';
import { ExceljsComponent } from './exceljs/exceljs.component';
import { ExcelDocumentResponseComponent } from './excel-document-response/excel-document-response.component';

const routes: Routes = [
  {path : 'basic', component: BasicComponent},
  {path : 'exceljs', component: ExceljsComponent},
  {path : 'ExcelDoc', component: ExcelDocumentResponseComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
