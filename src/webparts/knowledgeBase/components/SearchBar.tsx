import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';

interface ISearchBarProps {
  value: string;
  onChange: (value: string) => void;
}

const SearchBar: React.FC<ISearchBarProps> = ({ value, onChange }) => {
  return (
    <Stack.Item grow>
      <SearchBox
        placeholder="Search articles by title or description..."
        value={value}
        onChange={(_ev, newValue) => onChange(newValue || '')}
        onClear={() => onChange('')}
        styles={{
          root: { maxWidth: 400 }
        }}
        iconProps={{ iconName: 'Search' }}
      />
    </Stack.Item>
  );
};

export default SearchBar;
