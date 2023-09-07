import { Component, OnInit } from '@angular/core';
import { DataService } from 'src/app/_services/data.service';
import { ToastrService } from 'ngx-toastr';
import { formatDate } from '@angular/common';

@Component({
  selector: 'app-import-export',
  templateUrl: './import-export.component.html',
  styleUrls: ['./import-export.component.css'],
})
export class ImportExportComponent implements OnInit {
  fileList: File | null;

  constructor(private dataService: DataService, private toast: ToastrService) {}

  ngOnInit(): void {}

  onSelectedFile(event: Event) {
    const element = event.currentTarget as HTMLInputElement;
    this.fileList = element.files[0];
    console.log(this.fileList);
  }

  uploadFile(): void {
    if (!this.fileList) {
      this.toast.error('No file to upload');
    } else {
      const formData = new FormData();
      formData.append('formFile', this.fileList);
      this.dataService.import(formData).subscribe({
        next: () => this.toast.success('File uploaded successfully!'),
        error: (error) => console.log(error),
      });
    }
  }
  getFile() {
    this.dataService.export().subscribe({
      next: (response) => {
        const blob = new Blob([response], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        var fileLink = document.createElement('a');
        fileLink.href = url;
        fileLink.download =
          'Sample - ' + formatDate(new Date(), 'yyyy/MM/dd', 'en') + '.xlsx';
        fileLink.click();
      },
      error: (error) => console.log(error),
    });
  }
}
