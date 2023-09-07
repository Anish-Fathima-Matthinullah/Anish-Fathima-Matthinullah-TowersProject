import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ImportExportComponent } from './import-export/import-export.component';
import { DatatableExportComponent } from './datatable-export/datatable-export.component';

const routes: Routes = [
  { path: '', component: ImportExportComponent },
  { path: 'table', component: DatatableExportComponent },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule],
})
export class AppRoutingModule {}
