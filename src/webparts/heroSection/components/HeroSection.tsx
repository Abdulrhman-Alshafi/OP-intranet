// Trigger TS reload
import * as React from 'react';
import styles from './HeroSection.module.scss';
import type { IHeroSectionProps } from './IHeroSectionProps';

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

  // Determine grid columns based on image count
  const imgCount = imageList.length;
  const gridColumns = imgCount >= 2 ? '1fr 1fr' : '1fr';
  const gridRows = imgCount >= 3 ? '1fr 1fr' : undefined;

  return (
    <section className={styles.heroSection}>
      <div className={styles['heroLayout']}>
        {/* ── Left: Text ── */}
        <div className={styles['textContent']}>
          {heading && (
            <h1 className={styles.heading}>{heading}</h1>
          )}
          {description && (
            <p className={styles.description}>{description}</p>
          )}
        </div>

        {/* ── Right: Images ── */}
        <div
          className={styles['imagesContent']}
          style={{ gridTemplateColumns: gridColumns, gridTemplateRows: gridRows }}
        >
          {imageList.length > 0 ? (
            imageList.slice(0, 3).map((img, i) => (
              <div
                key={i}
                className={styles.imageWrapper}
                style={imgCount >= 3 && i === 0 ? { gridRow: '1 / 3' } : undefined}
              >
                <img src={img.image} alt={img.altText || `Hero image ${i + 1}`} />
              </div>
            ))
          ) : (
            <div className={styles.imagePlaceholder}>
              Add images via the property pane
            </div>
          )}
        </div>
      </div>
    </section>
  );
};

export default HeroSection;
