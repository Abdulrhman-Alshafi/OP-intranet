export interface ISlide {
  image?: string;
  headline?: string;
  description?: string;
  buttonText?: string;
  buttonLink?: string;
}

export interface IStaticTile {
  image?: string;
  bgColor?: string;
  title?: string;
  link?: string;
}

export interface ISwiperProps {
  slides: ISlide[];
  staticTiles: IStaticTile[];
  autoplayDelay: number;
  containerHeight: string;
  enablePagination: boolean;
  enableNavigation: boolean;
  buttonStyle: 'solid' | 'outline' | 'transparent';
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
