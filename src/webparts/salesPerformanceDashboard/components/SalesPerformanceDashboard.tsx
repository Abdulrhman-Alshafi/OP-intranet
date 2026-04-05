import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import styles from './SalesPerformanceDashboard.module.scss';
import { ISalesPerformanceDashboardProps } from './ISalesPerformanceDashboardProps';
import { ISalesData } from '../models/ISalesData';
import { SPService } from '../services/SPService';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { Dropdown, IDropdownOption, DocumentCard, Text, MessageBar, MessageBarType, Stack, IStackTokens, Spinner, SpinnerSize, useTheme } from '@fluentui/react';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';

const RC: any = ResponsiveContainer;
const BC: any = BarChart;
const CG: any = CartesianGrid;
const XA: any = XAxis;
const YA: any = YAxis;
const TT: any = Tooltip;
const LG: any = Legend;
const BR: any = Bar;

const stackTokens: IStackTokens = { childrenGap: 24 };

export default function SalesPerformanceDashboard(props: ISalesPerformanceDashboardProps) {
  const [data, setData] = useState<ISalesData[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedRegion, setSelectedRegion] = useState<string>('All');
  const [selectedProduct, setSelectedProduct] = useState<string>('All');

  const theme = useTheme();

  useEffect(() => {
    async function loadData() {
      if (!props.listId) {
        setLoading(false);
        return;
      }
      setLoading(true);
      const spData = await SPService.getSalesData(props.context, props.listId);
      
      if (spData && spData.length > 0) {
        setData(spData);
      } else {
        setData([]);
      }
      setLoading(false);
    }
    loadData().catch(console.error);
  }, [props.listId, props.context]);

  if (!props.listId) {
    return (
      <Placeholder 
        iconName="Edit"
        iconText="Configure Sales Dashboard"
        description="Please open the property pane and select a target SharePoint list to display data."
        buttonLabel="Configure"
        onConfigure={() => props.context.propertyPane.open()}
      />
    );
  }

  const uniqueRegions = useMemo(() => Array.from(new Set(data.map(d => d.Region))), [data]);
  const uniqueProducts = useMemo(() => Array.from(new Set(data.map(d => d.Product))), [data]);

  const filteredData = useMemo(() => {
    return data.filter(d => 
      (selectedRegion === 'All' || d.Region === selectedRegion) &&
      (selectedProduct === 'All' || d.Product === selectedProduct)
    );
  }, [data, selectedRegion, selectedProduct]);

  // Aggregate stats
  const totalTarget = useMemo(() => filteredData.reduce((sum, d) => sum + d.Target, 0), [filteredData]);
  const totalAchieved = useMemo(() => filteredData.reduce((sum, d) => sum + d.Achieved, 0), [filteredData]);
  const isTargetMissed = totalAchieved < totalTarget;

  // Chart data formatting
  const chartData = useMemo(() => {
    const grouped = filteredData.reduce((acc, curr) => {
      if (!acc[curr.Product]) {
        acc[curr.Product] = { name: curr.Product, Target: 0, Achieved: 0 };
      }
      acc[curr.Product].Target += curr.Target;
      acc[curr.Product].Achieved += curr.Achieved;
      return acc;
    }, {} as Record<string, any>);
    return Object.keys(grouped).map(key => grouped[key]);
  }, [filteredData]);

  const regionOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Regions' },
    ...uniqueRegions.map(r => ({ key: r, text: r }))
  ];

  const productOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Products' },
    ...uniqueProducts.map(p => ({ key: p, text: p }))
  ];

  return (
    <Stack tokens={stackTokens} className={styles.salesPerformanceDashboard}>
      <Text variant="xLargePlus" styles={{ root: { fontWeight: 600, color: props.isDarkTheme ? theme.palette.white : theme.palette.neutralPrimary } }}>
        Sales Performance Overview
      </Text>

      <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end" wrap>
        <Dropdown 
          label="Region Filter"
          options={regionOptions}
          selectedKey={selectedRegion}
          onChange={(_, opt) => opt && setSelectedRegion(opt.key as string)}
          styles={{ root: { minWidth: 220 } }}
        />
        <Dropdown 
          label="Product Filter"
          options={productOptions}
          selectedKey={selectedProduct}
          onChange={(_, opt) => opt && setSelectedProduct(opt.key as string)}
          styles={{ root: { minWidth: 220 } }}
        />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
        <DocumentCard styles={{ root: { padding: '24px 32px', minWidth: 250, borderTop: `4px solid ${theme.palette.themePrimary}` } }}>
          <Text variant="mediumPlus" styles={{ root: { color: theme.palette.neutralSecondary, fontWeight: 500 } }}>Total Sales Target</Text>
          <Text variant="superLarge" styles={{ root: { fontWeight: 700, marginTop: 12, display: 'block', color: theme.palette.neutralPrimary } }}>
            ${totalTarget.toLocaleString()}
          </Text>
        </DocumentCard>

        <DocumentCard styles={{ root: { padding: '24px 32px', minWidth: 250, borderTop: `4px solid ${isTargetMissed ? theme.palette.red : theme.palette.green}` } }}>
          <Text variant="mediumPlus" styles={{ root: { color: theme.palette.neutralSecondary, fontWeight: 500 } }}>Total Achieved</Text>
          <Text variant="superLarge" styles={{ root: { fontWeight: 700, marginTop: 12, display: 'block', color: isTargetMissed ? theme.palette.redDark : theme.palette.green } }}>
            ${totalAchieved.toLocaleString()}
          </Text>
        </DocumentCard>
      </Stack>

      {isTargetMissed && totalTarget > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
          Overall targets are currently missed for the selected filters.
        </MessageBar>
      )}

      {loading ? (
        <Spinner size={SpinnerSize.large} label="Loading sales data..." styles={{ root: { marginTop: 40 } }} />
      ) : chartData.length === 0 ? (
        <MessageBar messageBarType={MessageBarType.info}>
          No sales data found for the selected list or filters.
        </MessageBar>
      ) : (
        <DocumentCard styles={{ root: { padding: 24, width: '100%', maxWidth: 'none' } }}>
          <Text variant="large" styles={{ root: { marginBottom: 30, display: 'block', fontWeight: 600, color: theme.palette.neutralPrimary } }}>
            Performance by Product
          </Text>
          <div style={{ height: 350, width: '100%', paddingRight: 20 }}>
            <RC width="100%" height="100%">
              <BC data={chartData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                <CG strokeDasharray="3 3" vertical={false} stroke={theme.palette.neutralLight} />
                <XA dataKey="name" tick={{ fill: theme.palette.neutralSecondary }} axisLine={{ stroke: theme.palette.neutralQuaternary }} />
                <YA tick={{ fill: theme.palette.neutralSecondary }} axisLine={false} tickLine={false} />
                <TT contentStyle={{ borderRadius: 8, border: `1px solid ${theme.palette.neutralLight}` }} />
                <LG wrapperStyle={{ paddingTop: 20 }} />
                <BR dataKey="Target" fill={theme.palette.themeLight} radius={[4, 4, 0, 0]} />
                <BR dataKey="Achieved" fill={theme.palette.themePrimary} radius={[4, 4, 0, 0]} />
              </BC>
            </RC>
          </div>
        </DocumentCard>
      )}
    </Stack>
  );
}
