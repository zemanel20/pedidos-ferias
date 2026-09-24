import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context && !_sp) {
    // Inicializa o SP com o contexto do SPFx
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

// Função opcional para inicializar o SP fora do contexto do WebPart
export const initializeSP = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};