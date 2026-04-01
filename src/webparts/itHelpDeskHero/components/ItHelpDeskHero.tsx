import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './ItHelpDeskHero.module.scss';
import type { IItHelpDeskHeroProps, IHelpDeskButton, IHelpDeskTile } from './IItHelpDeskHeroProps';

// Helper: build tile grid class based on tile count
function getTileGridClass(count: number): string {
  if (count === 1) return styles.tiles1;
  if (count === 2) return styles.tiles2;
  if (count >= 3 && count < 4) return styles.tiles3;
  return styles.tiles4;
}

const ItHelpDeskHero = (props: IItHelpDeskHeroProps): React.ReactElement => {
  const {
    title,
    titleColor,
    description,
    backgroundColor,
    leftColumnWidthPercent,
    buttons,
    tiles,
  } = props;

  const leftWidth = `${leftColumnWidthPercent || '40'}%`;
  const displayTiles = (tiles || []).slice(0, 4); // cap at 4
  const displayButtons = buttons || [];

  const containerStyle: React.CSSProperties = {
    backgroundColor: backgroundColor || '#ffffff',
  };

  const leftColumnStyle: React.CSSProperties = {
    width: leftWidth,
  };

  const titleStyle: React.CSSProperties = {
    color: titleColor || '#0078d4',
  };

  return (
    <section className={styles.container} style={containerStyle}>
      {/* ── Left Column ── */}
      <div className={styles.leftColumn} style={leftColumnStyle}>
        {title && (
          <h2 className={styles.title} style={titleStyle}>
            {escape(title)}
          </h2>
        )}

        {description && (
          <p className={styles.description}>{escape(description)}</p>
        )}

        {displayButtons.length > 0 && (
          <div className={styles.buttonsWrapper}>
            {displayButtons.map((btn: IHelpDeskButton, idx: number) => {
              const linkProps = btn.url
                ? {
                    href: btn.url,
                    target: btn.openInNewTab ? '_blank' as const : '_self' as const,
                    rel: btn.openInNewTab ? 'noreferrer noopener' : undefined,
                  }
                : {};
              const Tag = btn.url ? 'a' : 'div';
              return (
                <Tag key={idx} className={styles.actionButton} {...linkProps}>
                  {btn.iconName && (
                    <span className={styles.buttonIcon}>
                      <Icon iconName={btn.iconName} />
                    </span>
                  )}
                  <span>{escape(btn.label || '')}</span>
                </Tag>
              );
            })}
          </div>
        )}
      </div>

      {/* ── Right Column (Tiles Mosaic) ── */}
      <div className={styles.rightColumn}>
        {displayTiles.length === 0 ? (
          <div className={styles.emptyState}>
            No tiles configured. Open the property pane to add up to 4 image tiles.
          </div>
        ) : (
          <div className={`${styles.tilesGrid} ${getTileGridClass(displayTiles.length)}`}>
            {displayTiles.map((tile: IHelpDeskTile, idx: number) => {
              const Tag = tile.linkUrl ? 'a' : 'div';
              const linkProps = tile.linkUrl
                ? {
                    href: tile.linkUrl,
                    target: tile.openInNewTab ? '_blank' as const : '_self' as const,
                    rel: tile.openInNewTab ? 'noreferrer noopener' : undefined,
                  }
                : {};

              const hasImage = tile.imageUrl && tile.imageUrl.trim().length > 0;

              return (
                <Tag
                  key={idx}
                  className={`${styles.tile}${!hasImage ? ` ${styles.tilePlaceholder}` : ''}`}
                  {...linkProps}
                  aria-label={tile.title}
                >
                  {hasImage && (
                    <img
                      className={styles.tileImage}
                      src={tile.imageUrl}
                      alt={tile.title || ''}
                    />
                  )}

                  <div className={styles.tileOverlay}>
                    {tile.title && (
                      <p className={styles.tileTitle}>{escape(tile.title)}</p>
                    )}
                    {tile.linkText && (
                      <span className={styles.tileLinkText}>{escape(tile.linkText)}</span>
                    )}
                  </div>
                </Tag>
              );
            })}
          </div>
        )}
      </div>
    </section>
  );
};

export default ItHelpDeskHero;
