import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { AadHttpClientFactory } from '@microsoft/sp-http';

export interface IPowerBiEmbedReportsProps {
  // description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
  userGroups?: string[];
  aadHttpClient: AadHttpClientFactory;
}
