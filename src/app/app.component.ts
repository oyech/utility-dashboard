import {
  Component,
  OnInit,
  AfterViewInit,
  ViewChild,
  ElementRef,
  OnDestroy,
  ChangeDetectorRef
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { UtilityService } from './services/utility.service';
import { BillingRecord, DashboardMetrics, UnitInfo } from './models/billing.model';

declare const Chart: any;
declare const lucide: any;

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html'
})
export class AppComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('usageChart') chartCanvas?: ElementRef<HTMLCanvasElement>;

  availablePeriods: string[] = [];
  selectedPeriods: string[] = [];
  selectedTower: string = '';
  searchQuery: string = '';
  searchResults: UnitInfo[] = [];
  showDropdown: boolean = false;
  isSearching: boolean = false;

  currentUnit: UnitInfo | null = null;
  records: BillingRecord[] = [];

  metrics: DashboardMetrics = {
    elecUsage: 0,
    elecCost: 0,
    waterUsage: 0,
    waterCost: 0,
    grandTotal: 0,
    periodCount: 0
  };

  isExporting: boolean = false;

  // Upload modal state
  uploadModalOpen: boolean = false;
  isUploading: boolean = false;
  uploadPeriod: string = '';
  uploadFile: File | null = null;
  uploadStatus = {
    visible: false,
    type: 'info' as 'info' | 'success' | 'error',
    message: ''
  };

  private chartInstance: any = null;
  private searchDebounceTimer: any = null;

  constructor(
    private utilityService: UtilityService,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.fetchPeriods();
  }

  ngAfterViewInit(): void {
    this.refreshIcons();
  }

  ngOnDestroy(): void {
    if (this.chartInstance) {
      this.chartInstance.destroy();
      this.chartInstance = null;
    }
  }

  refreshIcons(): void {
    setTimeout(() => {
      if (typeof lucide !== 'undefined' && lucide.createIcons) {
        lucide.createIcons();
      }
    }, 50);
  }

  fetchPeriods(): void {
    this.utilityService.getPeriods().subscribe({
      next: (periods) => {
        this.availablePeriods = periods;
        this.selectedPeriods = [...periods];
        if (this.currentUnit) {
          this.loadUsage();
        }
        this.refreshIcons();
        this.cdr.detectChanges();
      },
      error: (err) => console.error('Error fetching periods:', err)
    });
  }

  isPeriodSelected(period: string): boolean {
    return this.selectedPeriods.includes(period);
  }

  togglePeriod(period: string): void {
    const idx = this.selectedPeriods.indexOf(period);
    if (idx > -1) {
      this.selectedPeriods.splice(idx, 1);
    } else {
      this.selectedPeriods.push(period);
    }

    if (this.currentUnit) {
      this.loadUsage();
    }
  }

  onSearchInput(): void {
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
    }

    if (!this.searchQuery && !this.selectedTower) {
      this.searchResults = [];
      this.showDropdown = false;
      return;
    }

    this.searchDebounceTimer = setTimeout(() => {
      this.performSearch();
    }, 250);
  }

  onTowerChange(): void {
    this.performSearch();
  }

  performSearch(): void {
    this.isSearching = true;
    this.utilityService.searchUnits(this.searchQuery, this.selectedTower).subscribe({
      next: (units) => {
        this.searchResults = units;
        this.showDropdown = true;
        this.isSearching = false;
        this.refreshIcons();
        this.cdr.detectChanges();
      },
      error: (err) => {
        console.error('Error searching units:', err);
        this.searchResults = [];
        this.isSearching = false;
        this.cdr.detectChanges();
      }
    });
  }

  selectUnit(unitObj: UnitInfo): void {
    this.currentUnit = unitObj;
    this.searchQuery = `${unitObj.unit} - ${unitObj.owner_name || ''}`;
    this.showDropdown = false;
    this.searchResults = [];

    this.loadUsage();
    this.refreshIcons();
  }

  loadUsage(): void {
    if (!this.currentUnit || !this.currentUnit.unit) return;

    this.utilityService.getUsage(this.currentUnit.unit, this.selectedPeriods).subscribe({
      next: (records) => {
        this.records = records || [];
        this.calculateMetrics(this.records);
        this.refreshIcons();
        this.cdr.detectChanges();
        this.renderChart();
      },
      error: (err) => {
        console.error('Error loading usage:', err);
        this.records = [];
        this.calculateMetrics([]);
        this.cdr.detectChanges();
        this.renderChart();
      }
    });
  }

  calculateMetrics(records: BillingRecord[]): void {
    let totalElec = 0;
    let totalElecCost = 0;
    let totalWater = 0;
    let totalWaterCost = 0;
    let totalGrand = 0;

    for (const r of records) {
      totalElec += (r.elec_usage || 0);
      totalElecCost += (r.elec_total || 0);
      totalWater += (r.water_usage || 0);
      totalWaterCost += (r.water_total || 0);
      totalGrand += (r.grand_total || 0);
    }

    this.metrics = {
      elecUsage: totalElec,
      elecCost: totalElecCost,
      waterUsage: totalWater,
      waterCost: totalWaterCost,
      grandTotal: totalGrand,
      periodCount: records.length
    };
  }

  renderChart(): void {
    if (!this.chartCanvas) return;

    const ctx = this.chartCanvas.nativeElement.getContext('2d');
    if (!ctx) return;

    if (this.chartInstance) {
      this.chartInstance.destroy();
      this.chartInstance = null;
    }

    if (!this.records || this.records.length === 0) return;

    const labels = this.records.map(r => r.period);
    const elecData = this.records.map(r => r.elec_usage || 0);
    const waterData = this.records.map(r => r.water_usage || 0);

    this.chartInstance = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label: 'Listrik (kWh)',
            data: elecData,
            backgroundColor: 'rgba(245, 158, 11, 0.8)',
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            label: 'Air (m³)',
            data: waterData,
            backgroundColor: 'rgba(6, 182, 212, 0.8)',
            borderRadius: 6,
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            type: 'linear',
            position: 'left',
            title: { display: true, text: 'Listrik (kWh)' },
            grid: { color: 'rgba(0, 0, 0, 0.05)' }
          },
          y1: {
            type: 'linear',
            position: 'right',
            title: { display: true, text: 'Air (m³)' },
            grid: { drawOnChartArea: false }
          }
        }
      }
    });
  }

  exportToExcel(): void {
    if (!this.currentUnit || !this.currentUnit.unit) {
      alert('Silakan pilih unit terlebih dahulu sebelum export Excel.');
      return;
    }

    if (this.selectedPeriods.length === 0) {
      alert('Silakan pilih minimal 1 periode bulan untuk diekspor.');
      return;
    }

    this.isExporting = true;

    let chartImage: string | null = null;
    if (this.chartCanvas && this.chartCanvas.nativeElement) {
      try {
        chartImage = this.chartCanvas.nativeElement.toDataURL('image/png');
      } catch (e) {
        console.warn('Could not read chart image:', e);
      }
    }

    this.utilityService.exportExcel(this.currentUnit.unit, this.selectedPeriods, chartImage).subscribe({
      next: (blob) => {
        const filename = `Laporan_Utilitas_${this.currentUnit!.unit}.xlsx`;
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => window.URL.revokeObjectURL(url), 1000);
        this.isExporting = false;
        this.refreshIcons();
        this.cdr.detectChanges();
      },
      error: (err) => {
        console.error('Error exporting Excel:', err);
        alert('Terjadi kesalahan saat mengekspor file Excel.');
        this.isExporting = false;
        this.refreshIcons();
        this.cdr.detectChanges();
      }
    });
  }

  printSlip(): void {
    window.print();
  }

  openUploadModal(): void {
    this.uploadModalOpen = true;
    this.uploadPeriod = '';
    this.uploadFile = null;
    this.uploadStatus = { visible: false, type: 'info', message: '' };
    this.refreshIcons();
  }

  closeUploadModal(): void {
    this.uploadModalOpen = false;
    this.isUploading = false;
    this.uploadFile = null;
    this.uploadStatus = { visible: false, type: 'info', message: '' };
  }

  onFileSelected(event: any): void {
    const file = event.target.files?.[0];
    if (file) {
      this.uploadFile = file;
    }
  }

  submitUpload(): void {
    if (!this.uploadPeriod || !this.uploadFile) {
      this.uploadStatus = {
        visible: true,
        type: 'error',
        message: 'Mohon isi periode (YYYY-MM) dan pilih file .xlsx!'
      };
      return;
    }

    this.isUploading = true;
    this.uploadStatus = {
      visible: true,
      type: 'info',
      message: 'Sedang memproses dan mengimpor data...'
    };

    this.utilityService.uploadExcel(this.uploadPeriod, this.uploadFile).subscribe({
      next: (res) => {
        this.uploadStatus = {
          visible: true,
          type: 'success',
          message: res.message || 'Berhasil mengimpor data!'
        };
        this.cdr.detectChanges();
        setTimeout(() => {
          this.closeUploadModal();
          this.fetchPeriods();
          if (this.currentUnit) {
            this.loadUsage();
          }
          this.cdr.detectChanges();
        }, 1500);
      },
      error: (err) => {
        this.isUploading = false;
        this.uploadStatus = {
          visible: true,
          type: 'error',
          message: err.error?.error || 'Terjadi kesalahan saat upload.'
        };
        this.cdr.detectChanges();
      }
    });
  }

  formatRupiah(num: number | undefined | null): string {
    if (num === null || num === undefined || isNaN(num)) return 'Rp 0';
    return 'Rp ' + Math.round(Number(num)).toLocaleString('id-ID');
  }

  formatNumber(num: number | undefined | null, decimals: number = 1): string {
    if (num === null || num === undefined || isNaN(num)) return '0';
    return Number(num).toLocaleString('id-ID', {
      minimumFractionDigits: 0,
      maximumFractionDigits: decimals
    });
  }
}
