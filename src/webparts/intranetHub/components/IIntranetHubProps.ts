/* ── Shared item types ── */

export interface ILinkCardItem {
  iconName: string;
  title: string;
  linkUrl?: string;
}

export interface IDashboardStat {
  label: string;
  value: string;
}

export interface IDepartmentItem {
  iconName: string;
  title: string;
  linkUrl?: string;
  bgColor?: string;
}

/** Tools & Materials items carry a sectionTitle so they can be grouped under headings */
export interface IGroupedCardItem {
  sectionTitle: string;
  iconName: string;
  title: string;
  linkUrl?: string;
}

/* ── Main component props ── */

export interface IIntranetHubProps {
  /* Quick Links */
  showQuickLinks: boolean;
  quickLinksTitle: string;
  quickLinks: ILinkCardItem[];

  /* Dashboard */
  showDashboard: boolean;
  dashboardTitle: string;
  dashboardStats: IDashboardStat[];

  /* Departments */
  showDepartments: boolean;
  departmentsTitle: string;
  departments: IDepartmentItem[];

  /* Tools & Calculators */
  showToolsGrid: boolean;
  toolsItems: IGroupedCardItem[];

  /* Employee Spotlight */
  showSpotlight: boolean;
  spotlightName: string;
  spotlightRole: string;
  spotlightDescription: string;
  spotlightImageUrl: string;

  /* Materials & Resources */
  showMaterials: boolean;
  materialsItems: IGroupedCardItem[];

  /* Standard SPFx */
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
