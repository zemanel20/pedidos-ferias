import { IPedidos } from './IPedidos';

export class Pedidos implements IPedidos {
  public Id?: number;
  public Colaborador?: {
    Id: number;
    Title: string;
    EMail: string;
  };
  public Data_Inicio?: Date;
  public formattedData_Inicio?: string;
  public Data_Fim?: Date;
  public formattedData_Fim?: string;
  public Estado?: string;
}
