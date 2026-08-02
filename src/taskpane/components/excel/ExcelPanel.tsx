import React, { useState } from 'react';
import { makeStyles, tokens, TabList, Tab, Text } from '@fluentui/react-components';
import {
  DataAreaRegular, CalculatorRegular, DataBarVerticalRegular,
  MoneyRegular, BroomRegular, CodeRegular,
} from '@fluentui/react-icons';
import { DataTools } from './DataTools';
import { FormulaTools } from './FormulaTools';
import { ChartTools } from './ChartTools';
import { FinanceTools } from './FinanceTools';
import { CleaningTools } from './CleaningTools';
import { CodeGenTools } from './CodeGenTools';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
  },
  tabBar: {
    padding: '4px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    overflowX: 'auto',
    flexShrink: 0,
  },
  content: {
    flexGrow: 1,
    overflowY: 'auto',
    overflowX: 'hidden',
  },
});

type ToolCategory = 'analysis' | 'formula' | 'chart' | 'finance' | 'cleaning' | 'code';

const CATEGORIES = [
  { id: 'analysis' as ToolCategory, label: 'Analysis', icon: <DataAreaRegular /> },
  { id: 'formula' as ToolCategory, label: 'Formulas', icon: <CalculatorRegular /> },
  { id: 'chart' as ToolCategory, label: 'Charts', icon: <DataBarVerticalRegular /> },
  { id: 'finance' as ToolCategory, label: 'Finance', icon: <MoneyRegular /> },
  { id: 'cleaning' as ToolCategory, label: 'Cleaning', icon: <BroomRegular /> },
  { id: 'code' as ToolCategory, label: 'Code', icon: <CodeRegular /> },
];

export const ExcelPanel: React.FC = () => {
  const classes = useStyles();
  const [category, setCategory] = useState<ToolCategory>('analysis');

  const renderCategory = () => {
    switch (category) {
      case 'analysis': return <DataTools />;
      case 'formula': return <FormulaTools />;
      case 'chart': return <ChartTools />;
      case 'finance': return <FinanceTools />;
      case 'cleaning': return <CleaningTools />;
      case 'code': return <CodeGenTools />;
      default: return <DataTools />;
    }
  };

  return (
    <div className={classes.container}>
      <div className={classes.tabBar}>
        <TabList
          selectedValue={category}
          onTabSelect={(_, data) => setCategory(data.value as ToolCategory)}
          size="small"
        >
          {CATEGORIES.map(cat => (
            <Tab key={cat.id} value={cat.id} icon={cat.icon}>{cat.label}</Tab>
          ))}
        </TabList>
      </div>
      <div className={classes.content}>
        {renderCategory()}
      </div>
    </div>
  );
};
