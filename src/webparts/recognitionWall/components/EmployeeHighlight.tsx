import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IEmployeeHighlightProps } from './IRecognitionWallProps';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';

/**
 * EmployeeHighlight displays a banner for the Employee of the Month.
 * Shows their profile picture (via Fluent UI Persona), name, and info.
 * Admins can remove the selection via an X button.
 */
export default class EmployeeHighlight extends React.Component<IEmployeeHighlightProps> {

  public render(): React.ReactElement<IEmployeeHighlightProps> {
    const { employee, siteUrl, isAdmin, onRemove } = this.props;

    // Build profile picture URL from SharePoint user photos
    const email = 'email' in employee ? employee.email : ('EmployeeEmail' in employee ? (employee as any).EmployeeEmail : '');
    const name = 'name' in employee ? employee.name : ('EmployeeName' in employee ? (employee as any).EmployeeName : 'Unknown');
    const secondaryText = 'kudosCount' in employee ? `${employee.kudosCount} kudos this month` : 'Employee of the Month';

    const photoUrl = `${siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;

    return (
      <div className={styles.employeeHighlight}>
        <div className={styles.highlightBadge}>
          <Icon iconName="FavoriteStar" className={styles.starIcon} />
          <span>Employee of the Month</span>
        </div>
        <div className={styles.highlightContent}>
          <Persona
            imageUrl={photoUrl}
            text={name}
            secondaryText={secondaryText}
            size={PersonaSize.size72}
            imageAlt={`Profile picture of ${name}`}
          />
        </div>
        {isAdmin && onRemove && (
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            title="Remove Employee of the Month"
            ariaLabel="Remove Employee of the Month"
            className={styles.deleteBtn}
            onClick={onRemove}
          />
        )}
      </div>
    );
  }
}
