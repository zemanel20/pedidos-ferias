export interface IPedidos {
  Id?: number;
  Colaborador?: {
    Id: number;
    Title: string;
    EMail: string;
  };
  Data_Inicio?: Date;
  formattedData_Inicio?: string;
  Data_Fim?: Date;
  formattedData_Fim?: string;
  Estado?: string;
}
