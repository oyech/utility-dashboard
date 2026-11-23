import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { Observable } from 'rxjs';
import { BillingRecord, UnitInfo } from '../models/billing.model';

@Injectable({
  providedIn: 'root'
})
export class UtilityService {
  constructor(private http: HttpClient) {}

  getPeriods(): Observable<string[]> {
    return this.http.get<string[]>('/api/periods');
  }

  getTowers(): Observable<string[]> {
    return this.http.get<string[]>('/api/towers');
  }

  searchUnits(query: string, tower: string): Observable<UnitInfo[]> {
    let params = new HttpParams()
      .set('q', query || '')
      .set('tower', tower || '');
    return this.http.get<UnitInfo[]>('/api/units', { params });
  }

  getUsage(unit: string, periods: string[]): Observable<BillingRecord[]> {
    let params = new HttpParams().set('unit', unit);
    if (periods && periods.length > 0) {
      periods.forEach(p => {
        params = params.append('periods', p);
      });
    }
    return this.http.get<BillingRecord[]>('/api/usage', { params });
  }

  exportExcel(unit: string, periods: string[], chartImage?: string | null): Observable<Blob> {
    if (chartImage) {
      return this.http.post('/api/export-excel', {
        unit,
        periods: periods || [],
        chartImage
      }, {
        responseType: 'blob'
      });
    }

    let params = new HttpParams().set('unit', unit);
    if (periods && periods.length > 0) {
      periods.forEach(p => {
        params = params.append('periods', p);
      });
    }
    return this.http.get('/api/export-excel', {
      params,
      responseType: 'blob'
    });
  }

  uploadExcel(period: string, file: File): Observable<{ message: string }> {
    const formData = new FormData();
    formData.append('period', period);
    formData.append('file', file);
    return this.http.post<{ message: string }>('/api/upload', formData);
  }
}
