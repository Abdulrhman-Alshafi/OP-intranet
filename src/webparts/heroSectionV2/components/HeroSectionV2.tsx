import * as React from 'react';
// @ts-ignore - Ignore SPFx SCSS typings issue
import styles from './HeroSectionV2.module.scss';
import type { IHeroSectionV2Props } from './IHeroSectionV2Props';

const HeroSectionV2 = (props: IHeroSectionV2Props): React.ReactElement => {
  const {
    heading,
    subheading,
    backgroundImage,
    ctaText,
    ctaLink,
    showCta,
    badgeText,
    overlayColor,
    overlayOpacity,
    heroHeight,
    textColor,
    textAlignment
  } = props;

  const hasContent = heading || subheading || backgroundImage;

  if (!hasContent) {
    return (
      <section className={styles['heroSectionV2']}>
        <div className={styles['emptyState']}>
          <h2>Hero Section V2</h2>
          <p>Open the property pane to configure the heading, background image, and appearance.</p>
        </div>
      </section>
    );
  }

  // Build the overlay gradient
  const opacity = (overlayOpacity ?? 50) / 100;
  const color = overlayColor || '#000000';
  const overlayStyle: React.CSSProperties = {
    background: `linear-gradient(135deg, ${hexToRgba(color, opacity)} 0%, ${hexToRgba(color, opacity * 0.6)} 60%, ${hexToRgba(color, opacity * 0.3)} 100%)`
  };

  const containerStyle: React.CSSProperties = {
    height: `${heroHeight || 500}px`, // Enforce strict height
    minHeight: `${heroHeight || 500}px`,
    background: backgroundImage ? undefined : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)'
  };

  const textStyle: React.CSSProperties = {
    color: textColor || '#ffffff',
  };

  const contentLayerStyle: React.CSSProperties = {
    height: `${heroHeight || 500}px`,
    minHeight: `${heroHeight || 500}px`,
    textAlign: textAlignment || 'center',
    alignItems: textAlignment === 'left' ? 'flex-start' : textAlignment === 'right' ? 'flex-end' : 'center',
  };

  return (
    <section className={styles['heroSectionV2']} style={containerStyle}>
      {/* Background image */}
      {backgroundImage && (
        <img
          className={styles['backgroundImage']}
          src={backgroundImage}
          alt=""
          aria-hidden="true"
        />
      )}

      {/* Gradient overlay */}
      {backgroundImage && <div className={styles['overlay']} style={overlayStyle} />}

      {/* Content */}
      <div className={styles['contentLayer']} style={contentLayerStyle}>
        {badgeText && (
          <span className={styles['badge']}>{badgeText}</span>
        )}

        {heading && (
          <h1 className={styles['heading']} style={textStyle}>{heading}</h1>
        )}

        {subheading && (
          <p className={styles['subheading']} style={textStyle}>{subheading}</p>
        )}

        {showCta && ctaText && (
          <a
            className={styles['ctaButton']}
            href={ctaLink || '#'}
            target={ctaLink ? '_blank' : undefined}
            rel={ctaLink ? 'noopener noreferrer' : undefined}
          >
            {ctaText}
            <span className={styles['ctaArrow']}>→</span>
          </a>
        )}
      </div>
    </section>
  );
};

/**
 * Convert hex colour + alpha to rgba() string.
 */
function hexToRgba(hex: string, alpha: number): string {
  const cleaned = hex.replace('#', '');
  const r = parseInt(cleaned.substring(0, 2), 16) || 0;
  const g = parseInt(cleaned.substring(2, 4), 16) || 0;
  const b = parseInt(cleaned.substring(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

export default HeroSectionV2;
