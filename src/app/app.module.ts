import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { AppComponent } from './app.component';
import { FileService } from './services/file.service';
import { HttpClientModule } from '@angular/common/http';
import { UploadComponent } from './components/upload/upload.component';
import { DownloadComponent } from './components/download/download.component';
import { FileManagerComponent } from './components/file-manager/file-manager.component';
import {WatermarkManagerComponent} from './components/watermark-manager/watermark-manager.component';
import {WatermarkDownloadComponent} from './components/watermark-download/watermark-download.component';



@NgModule({
  declarations: [
    AppComponent,
    FileManagerComponent,
    UploadComponent,
    DownloadComponent,
    WatermarkManagerComponent,
    WatermarkDownloadComponent
  ],
  imports: [
    BrowserModule,
    HttpClientModule,
  ],
  providers: [FileService],
  bootstrap: [AppComponent]
})
export class AppModule { }
