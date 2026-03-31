import * as React from 'react';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import styles from './PersonCard.module.scss';

export interface IPersonCardProps {
  person: IPropertyFieldGroupOrPerson;
}

export const PersonCard: React.FC<IPersonCardProps> = ({ person }) => {
  // Use SharePoint's userphoto.aspx to resolve the profile image using the email
  const imageUrl = `/_layouts/15/userphoto.aspx?size=L&accountname=${person.email}`;

  return (
    <div className={styles.card}>
      <Persona
        imageUrl={imageUrl}
        text={person.fullName}
        secondaryText={person.email}
        size={PersonaSize.size72}
        presence={PersonaPresence.none}
      />
    </div>
  );
};
