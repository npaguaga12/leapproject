import { Component, Input} from '@angular/core';
import { HttpEventType } from '@angular/common/http';
import { FileService } from 'src/app/services/file.service';


@Component({
    selector: 'app-watermarkdownload',
    templateUrl: 'watermark-download.component.html'
})

export class WatermarkDownloadComponent{
    @Input() public disabled: boolean;
    @Input() public fileName: string;

    constructor(private service: FileService) {}

    public watermarkDownload(){
        this.service.watermarkDownloadFile(this.fileName).subscribe(
            data => {
                switch (data.type) {
                    case HttpEventType.Response:
                      const downloadedFile = new Blob([data.body], { type: data.body.type });
                      const a = document.createElement('a');
                      a.setAttribute('style', 'display:none;');
                      document.body.appendChild(a);
                      a.download = this.fileName;
                      a.href = URL.createObjectURL(downloadedFile);
                      a.target = '_blank';
                      a.click();
                      document.body.removeChild(a);
                      break;
                }
            }
        );
    }
}