import * as React from 'react';
import styles from './ServicesGrid.module.scss';
import type { IServicesGridProps } from './IServicesGridProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ServicesGrid extends React.Component<IServicesGridProps> {
  public render(): React.ReactElement<IServicesGridProps> {
    const {
      sectionTitle,
      services,
      columns
    } = this.props;

    const columnCount = parseInt(columns) || 4;
    const gridStyle = {
      gridTemplateColumns: `repeat(${columnCount}, 1fr)`
    };

    if (!services || services.length === 0) {
      return (
        <section className={styles.servicesGridContainer}>
          <div className={styles.emptyState}>
            <h2>{escape(sectionTitle)}</h2>
            <p>Please configure your cards by clicking 'Manage Cards' in the property pane.</p>
          </div>
        </section>
      );
    }

    return (
      <section className={styles.servicesGridContainer}>
        {sectionTitle && <h2 className={styles.sectionTitle}>{escape(sectionTitle)}</h2>}
        <div className={styles.grid} style={gridStyle}>
          {services.map((service, index) => {
            const hexColor = service.cardColor || 'transparent'; 
            const CardElement = service.linkUrl ? 'a' : 'div';
            const linkProps = service.linkUrl ? { href: service.linkUrl, target: '_blank', rel: 'noreferrer' } : {};

            return (
              <CardElement
                key={index}
                className={styles.serviceCard}
                style={{
                  cursor: service.linkUrl ? 'pointer' : 'default',
                  borderTopColor: hexColor !== 'transparent' ? hexColor : undefined
                }}
                {...linkProps}
              >
                {service.imageUrl && (
                  <div className={styles.imageWrapper}>
                    <img src={service.imageUrl} alt={service.title} />
                  </div>
                )}
                <h3 className={styles.title}>{escape(service.title || '')}</h3>
                <p className={styles.description}>{escape(service.description || '')}</p>
                {service.linkUrl && service.linkText && (
                  <div className={styles.learnMore} style={{ color: hexColor !== 'transparent' ? hexColor : '#0078d4' }}>
                    {escape(service.linkText)}
                  </div>
                )}
              </CardElement>
            );
          })}
        </div>
      </section>
    );
  }
}
