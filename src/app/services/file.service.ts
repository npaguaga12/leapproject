import { Injectable } from '@angular/core';
import { HttpClient, HttpRequest, HttpEvent, HttpResponse } from '@angular/common/http';
import { Observable } from 'rxjs';

@Injectable()
export class FileService{
    private baseApiUrl: string;
    private apiDownloadUrl: string;
    private apiUploadUrl: string;
    private apiFileUrl: string;
    private apiWatermarkFileUrl: string;
    private apiWatermarkDownloadUrl: string;

    constructor(private httpClient: HttpClient) {
        this.baseApiUrl = 'http://localhost:5000/api/';
        this.apiDownloadUrl = this.baseApiUrl + 'download';
        this.apiUploadUrl = this.baseApiUrl + 'upload';
        this.apiFileUrl = this.baseApiUrl + 'files';
        this.apiWatermarkFileUrl = this.baseApiUrl + 'watermarkfiles';
        this.apiWatermarkDownloadUrl = this.baseApiUrl + 'watermarkdownload';

    }

    public downloadFile(file: string): Observable<HttpEvent<Blob>> {
        return this.httpClient.request(new HttpRequest(
          'GET',
          `${this.apiDownloadUrl}?file=${file}`,
          null,
          {
            reportProgress: true,
            responseType: 'blob'
          }));
    }

    public watermarkDownloadFile(file: string): Observable<HttpEvent<Blob>> {
      return this.httpClient.request(new HttpRequest(
        'GET',
        `${this.apiWatermarkDownloadUrl}?file=${file}`));
    }

    public uploadFile(file: Blob): Observable<HttpEvent<void>> {
        const formData = new FormData();
        formData.append('file', file);
     
        return this.httpClient.request(new HttpRequest(
          'POST',
          this.apiUploadUrl,
          formData,
          {
            reportProgress: true
          }));
    }

    public getFiles(): Observable<string[]> {
        return this.httpClient.get<string[]>(this.apiFileUrl);
    }

    public getWatermarkFiles(): Observable<string[]> {
      return this.httpClient.get<string[]>(this.apiWatermarkFileUrl);
    }
    
}