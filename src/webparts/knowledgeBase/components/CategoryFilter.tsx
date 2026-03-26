import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';

interface ICategoryFilterProps {
  categories: string[];
  selectedCategory: string;
  onCategoryChange: (category: string) => void;
}

const CategoryFilter: React.FC<ICategoryFilterProps> = ({
  categories,
  selectedCategory,
  onCategoryChange
}) => {
  return (
    <Pivot
      selectedKey={selectedCategory}
      onLinkClick={(item) => {
        if (item?.props?.itemKey) {
          onCategoryChange(item.props.itemKey);
        }
      }}
      styles={{
        root: { marginBottom: 0 }
      }}
    >
      <PivotItem headerText="All" itemKey="All" />
      {categories.map((cat) => (
        <PivotItem key={cat} headerText={cat} itemKey={cat} />
      ))}
    </Pivot>
  );
};

export default CategoryFilter;
