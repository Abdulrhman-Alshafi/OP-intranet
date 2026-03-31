import * as React from 'react';
import styles from './OPTeamMembers.module.scss';
import { IOPTeamMembersProps } from './IOPTeamMembersProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PersonCard } from './PersonCard';

export default class OPTeamMembers extends React.Component<IOPTeamMembersProps, {}> {
  public render(): React.ReactElement<IOPTeamMembersProps> {
    const {
      title,
      people,
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.opTeamMembers} ${hasTeamsContext ? styles.teams : ''}`}>
        {title && <h2 className={styles.title}>{escape(title)}</h2>}
        
        {(!people || people.length === 0) ? (
          <div style={{ padding: '20px', textAlign: 'center', backgroundColor: 'var(--neutralLighter, #f3f2f1)' }}>
            Please edit the web part to select team members.
          </div>
        ) : (
          <div className={styles.grid}>
            {people.map((person, index) => (
              <PersonCard key={person.id || index} person={person} />
            ))}
          </div>
        )}
      </section>
    );
  }
}
