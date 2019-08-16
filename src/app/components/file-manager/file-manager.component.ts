import { Component, OnInit } from '@angular/core';
import { FileService } from 'src/app/services/file.service';
import { ProgressStatusEnum, ProgressStatus } from 'src/app/models/status.model';

@Component({
  selector: 'app-filemanager',
  templateUrl: './file-manager.component.html'
})
export class FileManagerComponent implements OnInit {
    public files: string[];
    public fileInDownload: string;
    public percentage: number;
    public showProgress: boolean;
    public showDownloadError: boolean;
    public showUploadError: boolean;

    constructor(private service: FileService) { }

    ngOnInit() {
      this.getFiles();
    }

    private getFiles() {
      this.service.getFiles().subscribe(
        data => {
          this.files = data;
          console.log(data);
        }
      );
    }

    public downloadStatus(event: ProgressStatus) {
        switch (event.status) {
          case ProgressStatusEnum.START:
            this.showDownloadError = false;
            break;
          case ProgressStatusEnum.IN_PROGRESS:
            this.showProgress = true;
            this.percentage = event.percentage;
            break;
          case ProgressStatusEnum.COMPLETE:
            this.showProgress = false;
            break;
          case ProgressStatusEnum.ERROR:
            this.showProgress = false;
            this.showDownloadError = true;
            break;
        }
    }
    public uploadStatus(event: ProgressStatus) {
        switch (event.status) {
          case ProgressStatusEnum.START:
            this.showUploadError = false;
            break;
          case ProgressStatusEnum.IN_PROGRESS:
            this.showProgress = true;
            this.percentage = event.percentage;
            break;
          case ProgressStatusEnum.COMPLETE:
            this.showProgress = false;
            this.getFiles();
            break;
          case ProgressStatusEnum.ERROR:
            this.showProgress = false;
            this.showUploadError = true;
            break;
        }
    }
}