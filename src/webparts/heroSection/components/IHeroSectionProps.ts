export interface IHeroImage {
  image: string;
  altText?: string;
}

export interface IHeroSectionProps {
  heading: string;
  description: string;
  images: IHeroImage[];
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
