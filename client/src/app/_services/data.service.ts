import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { environment } from 'src/environments/environment.development';
import { Api } from 'src/app/_models/api.model';

@Injectable({
  providedIn: 'root',
})
export class DataService {
  baseUrl = environment.apiUrl;

  constructor(private http: HttpClient) {}

  import(formData: FormData) {
    return this.http.post(this.baseUrl + 'Excel/import', formData);
  }

  export() {
    return this.http.get(this.baseUrl + 'Excel/export', {
      responseType: 'blob',
    });
  }

  getData() {
    return this.http.get<Api[]>(this.baseUrl + 'Excel');
  }

  updateData(payload: Api) {
    return this.http.put(this.baseUrl + 'Excel/update', payload);
  }
}
