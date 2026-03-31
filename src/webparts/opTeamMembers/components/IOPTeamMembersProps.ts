import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IOPTeamMembersProps {
  title: string;
  people: IPropertyFieldGroupOrPerson[];
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
