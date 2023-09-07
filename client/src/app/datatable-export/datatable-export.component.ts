import { Component, ViewChild, OnInit, ChangeDetectorRef } from '@angular/core';
import { ColumnMode, DatatableComponent } from '@swimlane/ngx-datatable';
import { DataService } from '../_services/data.service';
import { ToastrService } from 'ngx-toastr';
import { Api } from '../_models/api.model';

@Component({
  selector: 'app-datatable-export',
  templateUrl: './datatable-export.component.html',
  styleUrls: ['./datatable-export.component.css'],
})
export class DatatableExportComponent implements OnInit {
  title = 'agtable';
  ColumnMode = ColumnMode;
  editing = {};
  rows: Api[] = [];
  columns = [];
  percentages = [
    'activityComplete',
    'materialCostComplete',
    'laborCostComplete',
    'nonLaborCostComplete',
  ];

  constructor(
    private dataService: DataService,
    private toast: ToastrService,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    if (this.rows.length == 0) this.getRowData();
  }

  getRowClass = (row: any) => {
    let clrcode = row.color;
    return {
      'clr-beige': clrcode == '#FFFFF2CC',
      'clr-serenade': clrcode == '#FFFBE5D6',
      'clr-grey': clrcode == '#FFD9D9D9',
      'clr-green': clrcode == '#FF7DFFB8',
      'clr-blue': clrcode == '#FFDEEBF7',
      'clr-lgreen': clrcode == '#FFE2F0D9',
      'clr-purple': clrcode == '#FFDCC5ED',
    };
  };

  getTrimLength(str: string) {
    return str.length - str.trimStart().length;
  }
  getTrimedValue(str: string) {
    return str.trim();
  }

  getRowData() {
    this.dataService.getData().subscribe({
      next: (response: Api[]) => {
        this.rows = response;
        this.cdr.markForCheck();
      },
      error: (error) => this.toast.error('Unable to load!'),
    });
  }

  updateValue(event, cell, rowIndex, trimLength) {
    this.editing[`${rowIndex}${cell}`] = false;
    if (cell === 'activityId') {
      this.rows[rowIndex][cell] = event.target.value.padStart(
        event.target.value.length + trimLength,
        ' '
      );
    } else if (this.percentages.includes(cell)) {
      this.rows[rowIndex][cell] = event.target.value / 100;
    } else {
      this.rows[rowIndex][cell] = event.target.value;
    }

    this.rows = [...this.rows];
    let payload: Api = this.rows[rowIndex];

    this.dataService.updateData(payload).subscribe({
      next: () => {
        this.getRowData();
        this.toast.success('Saved');
      },
      error: (error) => {
        this.toast.warning('Cancelled!');
        this.getRowData();
      },
    });
  }
}
