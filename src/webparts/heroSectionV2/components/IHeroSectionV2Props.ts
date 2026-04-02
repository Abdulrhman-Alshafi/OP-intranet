export interface IHeroSectionV2Props {
  heading: string;
  subheading: string;
  backgroundImage: string;
  ctaText: string;
  ctaLink: string;
  showCta: boolean;
  badgeText: string;
  overlayColor: string;
  overlayOpacity: number;
  heroHeight: number;
  textColor: string;
  textAlignment: 'left' | 'center' | 'right';
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
