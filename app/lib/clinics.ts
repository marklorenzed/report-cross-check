export const MANUAL_REPORT_CLINICS = [
  { id: "staRosa", label: "DMP Dental Clinic Sta. Rosa" },
  { id: "calamba", label: "DMP Dental Clinic Calamba" },
  { id: "sanPedro", label: "DMP Dental Clinic San Pedro" },
  { id: "stoDomingo", label: "DMP Dental Clinic Sto. Domingo" },
] as const;

export type ManualReportClinicId = (typeof MANUAL_REPORT_CLINICS)[number]["id"];
