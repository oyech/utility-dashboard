export interface UnitInfo {
  unit: string;
  tower: string;
  owner_name: string;
  kva?: number;
  mcb?: number;
}

export interface BillingRecord {
  id?: number;
  period: string;
  no?: number;
  tower: string;
  unit: string;
  owner_name: string;
  mcb?: number;
  kva?: number;
  note?: string;
  elec_meter_start?: number;
  elec_meter_end?: number;
  elec_usage?: number;
  kwh_min?: number;
  jam_nyala?: number;
  elec_subtotal?: number;
  elec_total?: number;
  water_meter_start?: number;
  water_meter_end?: number;
  water_usage?: number;
  water_m1?: number;
  water_m2?: number;
  water_m3?: number;
  water_cost_m1?: number;
  water_cost_m2?: number;
  water_cost_m3?: number;
  water_fixed_cost?: number;
  water_total?: number;
  grand_total?: number;
}

export interface DashboardMetrics {
  elecUsage: number;
  elecCost: number;
  waterUsage: number;
  waterCost: number;
  grandTotal: number;
  periodCount: number;
}
