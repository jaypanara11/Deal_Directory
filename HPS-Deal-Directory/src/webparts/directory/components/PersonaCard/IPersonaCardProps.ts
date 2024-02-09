import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IUserProperties } from "./IUserProperties";
import { ITickerProperties } from "../../../../SPServices/ITickerProperties";

export interface IPersonaCardProps {
  context: WebPartContext | ApplicationCustomizerContext;
  // profileProperties: IUserProperties;
  tickerProperties: ITickerProperties;
}
