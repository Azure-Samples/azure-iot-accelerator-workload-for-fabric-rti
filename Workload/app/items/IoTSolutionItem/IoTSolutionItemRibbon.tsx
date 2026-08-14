// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { PageProps } from "../../App";
import {
  Ribbon,
  RibbonAction,
  useSaveAction,
  useSettingsAction,
} from "../../components/ItemEditor";
import { ViewContext } from "../../components";

export interface IoTSolutionItemRibbonProps extends PageProps {
  isSaveButtonEnabled?: boolean;
  viewContext: ViewContext;
  saveItemCallback: () => Promise<void>;
  openSettingsCallback: () => Promise<void>;
}

/**
 * Ribbon for the IoT Solution item editor.
 * Provides Save and Settings actions.
 */
export function IoTSolutionItemRibbon(props: IoTSolutionItemRibbonProps) {
  const { viewContext } = props;

  const saveAction = useSaveAction(
    props.saveItemCallback,
    !props.isSaveButtonEnabled
  );

  const settingsAction = useSettingsAction(props.openSettingsCallback);

  const homeToolbarActions: RibbonAction[] = [saveAction, settingsAction];

  return (
    <Ribbon
      homeToolbarActions={homeToolbarActions}
      additionalToolbars={[]}
      rightActionButtons={[]}
      viewContext={viewContext}
    />
  );
}
