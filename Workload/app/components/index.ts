// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Reusable Components for Microsoft Fabric Workload Items
 * 
 * This module exports commonly used UI components that maintain consistency
 * across all item editors in the workload.
 * 
 * @see {@link ../../docs/components/README.md} - Complete components documentation overview
 * @see {@link ../../docs/components/ItemEditor.md} - ItemEditor component and architecture
 * @see {@link ../../docs/components/Wizard.md} - Wizard component for step-by-step workflows
 */

// Base Item Editor - Foundation for all item editors
export { 
  ItemEditor,
  ItemEditorDefaultView,
  ItemEditorEmptyView,
  ItemEditorDetailView,
  Ribbon,
  RibbonToolbar,
  RibbonToolbarAction,
  useSaveAction,
  useSettingsAction,
  useAboutAction
} from './ItemEditor/';

export type { 
  ItemEditorProps, 
  RegisteredView,
  ViewContext,
  LeftPanelConfig,
  CentralPanelConfig,
  EmptyStateTask,
  DetailViewAction,
  RibbonProps,
  RibbonAction,
  RibbonActionButton,
  FluentIconComponent
} from './ItemEditor/';

// Wizard Component - Step-by-step guided workflows
export { 
  WizardControl
} from './Wizard/';

export type { 
  WizardStep, 
  WizardControlProps,
  WizardStepProps,
  WizardNavigationProps
} from './Wizard/';

// Dialog Component - Simple dialog with wizard-like styling
export {
  DialogControl
} from './Dialog/';

export type {
  DialogControlProps
} from './Dialog/';
