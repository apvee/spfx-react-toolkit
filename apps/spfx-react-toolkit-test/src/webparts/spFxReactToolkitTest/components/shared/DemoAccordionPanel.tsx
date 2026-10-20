import * as React from 'react';
import { AccordionPanel, AccordionPanelProps } from '@fluentui/react-components';

// React 17 omits inert from its HTML types and requires the native empty attribute.
type InertPanelProps = React.HTMLAttributes<HTMLDivElement> & { inert?: '' };

// AccordionGroup owns card/header composition; diagnostics need only this native panel.
/** Retained Fluent diagnostics with the same React 17 motion adaptation as AccordionGroup. */
export const DemoAccordionPanel: React.FC<Omit<AccordionPanelProps, 'collapseMotion'>> = props => (
  <AccordionPanel {...props} collapseMotion={{
    children: (Motion, motionProps) => {
      const panel = React.isValidElement<InertPanelProps>(motionProps.children)
        ? React.cloneElement(motionProps.children, { inert: motionProps.visible ? undefined : '' })
        : motionProps.children;
      return <Motion {...motionProps} unmountOnExit={false}>{panel}</Motion>;
    },
  }} />
);
