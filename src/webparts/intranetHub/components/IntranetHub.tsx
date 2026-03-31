import * as React from 'react';
import styles from './IntranetHub.module.scss';
import type {
  IIntranetHubProps,
  ILinkCardItem,
  IDashboardStat,
  IDepartmentItem,
  IGroupedCardItem
} from './IIntranetHubProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { escape } from '@microsoft/sp-lodash-subset';

/* ────────────────────────────────────────
   Section renderers
   ──────────────────────────────────────── */

const renderQuickLinks = (title: string, items: ILinkCardItem[]): React.ReactElement => (
  <div>
    <h2 className={styles.sectionTitle}>{escape(title)}</h2>
    <div className={styles.quickLinksGrid}>
      {items.map((item, i) => {
        const Tag = item.linkUrl ? 'a' : 'div';
        const linkProps = item.linkUrl
          ? { href: item.linkUrl, target: '_blank' as const, rel: 'noreferrer' }
          : {};
        return (
          <Tag key={i} className={styles.linkCard} {...linkProps}>
            <Icon iconName={item.iconName || 'Link'} className={styles.linkCardIcon} />
            <span className={styles.linkCardTitle}>{escape(item.title || '')}</span>
          </Tag>
        );
      })}
    </div>
  </div>
);

const renderDashboard = (title: string, stats: IDashboardStat[]): React.ReactElement => (
  <div>
    <h2 className={styles.sectionTitle}>{escape(title)}</h2>
    <div className={styles.statsList}>
      {stats.map((stat, i) => (
        <div key={i} className={styles.statRow}>
          <span className={styles.statLabel}>{escape(stat.label || '')}</span>
          <span className={styles.statValue}>{escape(stat.value || '')}</span>
        </div>
      ))}
    </div>
  </div>
);

const renderDepartments = (title: string, items: IDepartmentItem[]): React.ReactElement => (
  <div>
    <h2 className={styles.sectionTitle}>{escape(title)}</h2>
    <div className={styles.departmentsGrid}>
      {items.map((item, i) => {
        const Tag = item.linkUrl ? 'a' : 'div';
        const linkProps = item.linkUrl
          ? { href: item.linkUrl, target: '_blank' as const, rel: 'noreferrer' }
          : {};
        const bg = item.bgColor || '#1B2A4A';
        return (
          <Tag key={i} className={styles.deptCard} style={{ backgroundColor: bg }} {...linkProps}>
            <Icon iconName={item.iconName || 'Org'} className={styles.deptIcon} />
            <span className={styles.deptTitle}>{escape(item.title || '')}</span>
          </Tag>
        );
      })}
    </div>
  </div>
);

/** Groups items by sectionTitle and renders a heading + 3-col grid per group */
const renderGroupedGrid = (items: IGroupedCardItem[]): React.ReactElement => {
  // Preserve insertion order via Map
  const groups = new Map<string, IGroupedCardItem[]>();
  items.forEach(item => {
    const key = item.sectionTitle || 'Untitled';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(item);
  });

  return (
    <div>
      {Array.from(groups.entries()).map(([groupTitle, groupItems], gi) => (
        <div key={gi} className={styles.groupedSection}>
          <h2 className={styles.groupTitle}>{escape(groupTitle)}</h2>
          <div className={styles.groupedGrid}>
            {groupItems.map((item, i) => {
              const Tag = item.linkUrl ? 'a' : 'div';
              const linkProps = item.linkUrl
                ? { href: item.linkUrl, target: '_blank' as const, rel: 'noreferrer' }
                : {};
              return (
                <Tag key={i} className={styles.toolCard} {...linkProps}>
                  <div className={styles.toolIcon}>
                    <Icon iconName={item.iconName || 'Settings'} />
                  </div>
                  <span className={styles.toolTitle}>{escape(item.title || '')}</span>
                </Tag>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
};

const renderSpotlight = (
  name: string,
  role: string,
  description: string,
  imageUrl: string
): React.ReactElement => (
  <div>
    <h2 className={styles.sectionTitle}>Employee Spotlight</h2>
    <div className={styles.spotlightCard}>
      {imageUrl ? (
        <img src={imageUrl} alt={name} className={styles.spotlightImage} />
      ) : (
        <div className={styles.spotlightPlaceholder}>Upload employee photo</div>
      )}
      <div className={styles.spotlightBody}>
        {name && <h3 className={styles.spotlightName}>{escape(name)}</h3>}
        {role && <p className={styles.spotlightRole}>{escape(role)}</p>}
        {description && <p className={styles.spotlightDesc}>{escape(description)}</p>}
      </div>
    </div>
  </div>
);

/* ────────────────────────────────────────
   Main component
   ──────────────────────────────────────── */

const IntranetHub = (props: IIntranetHubProps): React.ReactElement => {
  const {
    showQuickLinks, quickLinksTitle, quickLinks,
    showDashboard, dashboardTitle, dashboardStats,
    showDepartments, departmentsTitle, departments,
    showToolsGrid, toolsItems,
    showSpotlight, spotlightName, spotlightRole, spotlightDescription, spotlightImageUrl,
    showMaterials, materialsItems
  } = props;

  const hasAnySection =
    showQuickLinks || showDashboard || showDepartments ||
    showToolsGrid || showSpotlight || showMaterials;

  if (!hasAnySection) {
    return (
      <section className={styles.intranetHub}>
        <div className={styles.emptyState}>
          <h2>Intranet Hub</h2>
          <p>Open the property pane to enable and configure sections.</p>
        </div>
      </section>
    );
  }

  /* Top 3-column area: Quick Links | Dashboard | Departments */
  const hasTopRow = showQuickLinks || showDashboard || showDepartments;
  /* Bottom area: Tools + Materials (left 2/3) | Spotlight (right 1/3) */
  const hasBottomLeft = (showToolsGrid && toolsItems && toolsItems.length > 0) ||
                        (showMaterials && materialsItems && materialsItems.length > 0);
  const hasBottomRight = showSpotlight;
  const hasBottomRow = hasBottomLeft || hasBottomRight;

  return (
    <section className={styles.intranetHub}>
      {hasTopRow && (
        <div className={styles.masterGrid}>
          {showQuickLinks && renderQuickLinks(quickLinksTitle || 'Quick Links', quickLinks || [])}
          {showDashboard && renderDashboard(dashboardTitle || 'Dashboard', dashboardStats || [])}
          {showDepartments && renderDepartments(departmentsTitle || 'Departments', departments || [])}
        </div>
      )}

      {hasTopRow && hasBottomRow && <hr className={styles.divider} />}

      {hasBottomRow && (
        <div className={styles.bottomRow}>
          <div>
            {showToolsGrid && toolsItems && toolsItems.length > 0 && renderGroupedGrid(toolsItems)}
            {showMaterials && materialsItems && materialsItems.length > 0 && renderGroupedGrid(materialsItems)}
          </div>
          <div>
            {hasBottomRight && renderSpotlight(
              spotlightName || '',
              spotlightRole || '',
              spotlightDescription || '',
              spotlightImageUrl || ''
            )}
          </div>
        </div>
      )}
    </section>
  );
};

export default IntranetHub;
