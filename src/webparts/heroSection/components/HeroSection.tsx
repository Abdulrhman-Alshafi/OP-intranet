import * as React from 'react';
import styles from './HeroSection.module.scss';
import type { IHeroSectionProps } from './IHeroSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';

const HeroSection = (props: IHeroSectionProps): React.ReactElement => {
  const { heading, description, images } = props;
  const imageList = images || [];

  const hasContent = heading || description || imageList.length > 0;

  if (!hasContent) {
    return (
      <section className={styles.heroSection}>
        <div className={styles.emptyState}>
          <h2>Hero Section</h2>
          <p>Open the property pane to configure the heading, description, and images.</p>
        </div>
      </section>
    );
  }

  return (
    <section className={styles.heroSection}>
      {/* ── Top row: heading + description ── */}
      <div className={styles.topRow}>
        {heading && (
          <h1 className={styles.heading}>{escape(heading)}</h1>
        )}
        {description && (
          <p className={styles.description}>{escape(description)}</p>
        )}
      </div>

      {/* ── Bottom row: images ── */}
      {imageList.length > 0 && (
        <div className={styles.imagesRow}>
          {imageList.map((img, i) => (
            <div key={i} className={styles.imageWrapper}>
              <img src={img.image} alt={img.altText || `Hero image ${i + 1}`} />
            </div>
          ))}
        </div>
      )}

      {imageList.length === 0 && (
        <div className={styles.imagesRow}>
          <div className={styles.imagePlaceholder}>
            Add images via the property pane
          </div>
        </div>
      )}
    </section>
  );
};

export default HeroSection;
