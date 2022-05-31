import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IPowerBiEmbedReportsProps {
  // description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
  userGroups?: string[];
  accessToken?: string;
  accessTokenError?: string;
}
