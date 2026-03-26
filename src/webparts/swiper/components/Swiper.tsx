import * as React from 'react';
import importedStyles from './Swiper.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import type { ISwiperProps } from './ISwiperProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Swiper as SwiperReact, SwiperSlide } from 'swiper/react';
import { Autoplay, Pagination, Navigation, EffectFade } from 'swiper/modules';

import 'swiper/css';
import 'swiper/css/pagination';
import 'swiper/css/navigation';
import 'swiper/css/effect-fade';

const Swiper = (props: ISwiperProps): React.ReactElement => {
  const {
    slides,
    staticTiles,
    autoplayDelay,
    containerHeight,
    enablePagination,
    enableNavigation,
    buttonStyle,
    hasTeamsContext
  } = props;

  const hasStaticTiles = staticTiles && staticTiles.length > 0;
  const computedHeight = containerHeight && !isNaN(Number(containerHeight)) ? `${containerHeight}px` : '500px';
  const rootStyle = { '--hero-height': computedHeight } as React.CSSProperties;

  if (!slides || slides.length === 0) {
    return (
      <section className={`${styles.swiperContainer} ${hasTeamsContext ? styles.teams : ''}`} style={rootStyle}>
        <div className={styles.emptyState}>
          <h2>Welcome to the Swiper Web Part</h2>
          <p>Configure slides by clicking &lsquo;Manage Slides&rsquo; in the property pane.</p>
        </div>
      </section>
    );
  }

  const swiperContent = (
    <SwiperReact
      spaceBetween={0}
      slidesPerView={1}
      autoplay={autoplayDelay > 0 ? { delay: autoplayDelay, disableOnInteraction: false, pauseOnMouseEnter: true } : false}
      pagination={enablePagination ? { clickable: true } : false}
      navigation={enableNavigation}
      effect="fade"
      speed={500}
      modules={[Autoplay, Pagination, Navigation, EffectFade]}
      className={styles.swiperRoot}
    >
      {slides.map((slide, index) => (
        <SwiperSlide key={index}>
          <div
            className={styles.slideWrapper}
            style={{ backgroundImage: `url(${slide.image || ''})` }}
          >
            <div className={styles.slideOverlay}>
              <div className={styles.slideContent}>
                {slide.headline && <h2>{escape(slide.headline)}</h2>}
                {slide.description && <p>{escape(slide.description)}</p>}
                {slide.buttonText && slide.buttonLink && (
                  <a href={slide.buttonLink}
                     className={`${styles.ctaButton} ${
                       buttonStyle === 'outline' ? 'swiperWebPart_ctaButtonOutline' :
                       buttonStyle === 'transparent' ? 'swiperWebPart_ctaButtonTransparent' :
                       'swiperWebPart_ctaButtonSolid'
                     }`}
                     target="_blank" rel="noreferrer">
                    {escape(slide.buttonText)}
                  </a>
                )}
              </div>
            </div>
          </div>
        </SwiperSlide>
      ))}
    </SwiperReact>
  );

  if (hasStaticTiles) {
    return (
      <section className={`${styles.heroContainer} ${hasTeamsContext ? styles.teams : ''}`} style={rootStyle}>
        <div className={styles.heroMain}>
          {swiperContent}
        </div>
        <div className={styles.heroSide}>
          {staticTiles.map((tile, index) => {
            const bgStyle = tile.image
              ? { backgroundImage: `url(${tile.image})`, backgroundColor: tile.bgColor || '#0078d4' }
              : { backgroundColor: tile.bgColor || '#0078d4' };

            const content = (
              <>
                <div className={styles.tileOverlay} />
                <div className={styles.tileContent}>
                  {tile.title && <h3>{escape(tile.title)}</h3>}
                </div>
              </>
            );

            return tile.link ? (
              <a
                key={index}
                href={tile.link}
                className={styles.heroTile}
                style={bgStyle}
                target="_blank"
                rel="noreferrer"
              >
                {content}
              </a>
            ) : (
              <div
                key={index}
                className={styles.heroTile}
                style={bgStyle}
              >
                {content}
              </div>
            );
          })}
        </div>
      </section>
    );
  }

  return (
    <section className={`${styles.swiperContainer} ${hasTeamsContext ? styles.teams : ''}`} style={rootStyle}>
      {swiperContent}
    </section>
  );
};

export default Swiper;
