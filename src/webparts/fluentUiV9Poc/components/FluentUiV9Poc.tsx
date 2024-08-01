import * as React from 'react';
import styles from './FluentUiV9Poc.module.scss';
import type { IFluentUiV9PocProps } from './IFluentUiV9PocProps'; 
import { ControllingOpenAndClose } from './ControllingOpenAndClose';
import BreadCrumb from './BreadCrumb';
import ThemeWrapper from './ThemeWrapper';
import { Context } from './context';

const FluentUiV9Poc: React.FC<IFluentUiV9PocProps> = () => {
  const { hasTeamsContext } = React.useContext(Context);
    return (
      <ThemeWrapper >
        <section className={`${styles.fluentUiV9Poc} ${hasTeamsContext ? styles.teams : ''}`}>
          <BreadCrumb />
          <ControllingOpenAndClose /> 
        </section>
      </ThemeWrapper> 
    );
}

export default FluentUiV9Poc;
