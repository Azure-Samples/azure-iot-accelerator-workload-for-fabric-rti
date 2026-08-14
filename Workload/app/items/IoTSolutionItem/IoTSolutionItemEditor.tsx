// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useEffect, useState } from "react";
import { useParams, useLocation } from "react-router-dom";
import { NotificationType } from "@ms-fabric/workload-client";
import { Spinner } from "@fluentui/react-components";
import { PageProps, ContextProps } from "../../App";
import {
  ItemWithDefinition,
  getWorkloadItem,
  callGetItem,
  saveWorkloadItem,
} from "../../controller/ItemCRUDController";
import { callOpenSettings } from "../../controller/SettingsController";
import { callNotificationOpen } from "../../controller/NotificationController";
import { navigateToWorkspace } from "../../controller/NavigationController";
import {
  ItemEditor,
  useViewNavigation,
} from "../../components/ItemEditor";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { IoTSolutionItemHomePage } from "./IoTSolutionItemHomePage";
import { IoTSolutionItemDefaultView } from "./IoTSolutionItemDefaultView";
import { IoTSolutionItemDashboardWizard } from "./IoTSolutionItemDashboardWizard";
import { IoTSolutionItemModelWizard } from "./IoTSolutionItemModelWizard";
import { IoTSolutionItemActivatorWizard } from "./IoTSolutionItemActivatorWizard";
import { IoTSolutionItemDataAgentWizard } from "./IoTSolutionItemDataAgentWizard";
import { IoTSolutionItemRibbon } from "./IoTSolutionItemRibbon";
import "./IoTSolutionItem.scss";

export const EDITOR_VIEW_TYPES = {
  HOME: "home",
  INGESTION_WIZARD: "ingestion-wizard",
  DASHBOARD_WIZARD: "dashboard-wizard",
  MODEL_WIZARD: "model-wizard",
  ACTIVATOR_WIZARD: "activator-wizard",
  DATA_AGENT_WIZARD: "data-agent-wizard",
} as const;

const enum SaveStatus {
  NotSaved = "NotSaved",
  Saving = "Saving",
  Saved = "Saved",
}

export function IoTSolutionItemEditor(props: PageProps) {
  const { workloadClient } = props;
  const pageContext = useParams<ContextProps>();

  const [isLoading, setIsLoading] = useState(true);
  const [isSavingWizard, setIsSavingWizard] = useState(false);
  const [item, setItem] = useState<ItemWithDefinition<IoTSolutionItemDefinition>>();
  const [saveStatus, setSaveStatus] = useState<SaveStatus>(SaveStatus.NotSaved);
  const [currentDefinition, setCurrentDefinition] = useState<IoTSolutionItemDefinition>({});
  const [viewSetter, setViewSetter] = useState<((view: string) => void) | null>(null);

  const { pathname } = useLocation();

  async function loadDataFromUrl(ctx: ContextProps, path: string): Promise<void> {
    if (ctx.itemObjectId && item && item.id === ctx.itemObjectId) {
      return;
    }

    setIsLoading(true);
    if (ctx.itemObjectId) {
      try {
        let loadedItem = await getWorkloadItem<IoTSolutionItemDefinition>(
          workloadClient,
          ctx.itemObjectId
        );

        if (!loadedItem.definition) {
          setSaveStatus(SaveStatus.NotSaved);
          loadedItem = { ...loadedItem, definition: {} };
        } else {
          setSaveStatus(SaveStatus.Saved);
        }

        setItem(loadedItem);
        setCurrentDefinition(loadedItem.definition || {});
      } catch {
        setItem(undefined);
      }
    }
    setIsLoading(false);
  }

  useEffect(() => {
    loadDataFromUrl(pageContext, pathname);
  }, [pageContext, pathname]);

  const handleOpenSettings = async () => {
    if (item) {
      try {
        const itemRes = await callGetItem(workloadClient, item.id);
        await callOpenSettings(workloadClient, itemRes.item, "About");
      } catch (error) {
        console.error("Failed to open settings:", error);
      }
    }
  };

  async function saveItem() {
    setSaveStatus(SaveStatus.Saving);

    let successResult;
    let errorMessage = "";

    try {
      successResult = await saveWorkloadItem<IoTSolutionItemDefinition>(
        workloadClient,
        { ...item, definition: currentDefinition }
      );
    } catch (error: any) {
      errorMessage = error?.message;
    }

    if (successResult) {
      if (item) {
        item.definition = currentDefinition;
      }
      setSaveStatus(SaveStatus.Saved);
      callNotificationOpen(
        workloadClient,
        "Saved",
        `IoT Solution "${item?.displayName}" saved successfully.`,
        undefined,
        undefined
      );
    } else {
      setSaveStatus(SaveStatus.NotSaved);
      callNotificationOpen(
        workloadClient,
        "Save Failed",
        errorMessage || "Failed to save IoT Solution.",
        NotificationType.Error,
        undefined
      );
    }
  }

  const handleWizardComplete = async (
    wizardContext: Record<string, any>,
    navigateHome?: () => void
  ): Promise<void> => {
    // Show full-page saving overlay immediately
    setIsSavingWizard(true);

    // Determine which wizard completed based on context keys
    const isDashboardWizard = !!wizardContext.dashDashboardCreated;
    const isModelWizard = !!wizardContext.modelUploaded;
    const isActivatorWizard = !!wizardContext.activatorCreated;
    const isDataAgentWizard = !!wizardContext.agentCreated;

    // Map wizard context to item definition
    const updatedDef: IoTSolutionItemDefinition = {
      ...currentDefinition,
      wizardStep: "completed",
      // Ingestion wizard fields (only update if coming from ingestion wizard)
      ...(wizardContext.iotHubValidated && {
        iotHub: wizardContext.iotHubDetails
          ? {
              subscriptionId: wizardContext.subscriptionId,
              resourceGroup: wizardContext.resourceGroup,
              hubName: wizardContext.hubName,
              location: wizardContext.iotHubDetails.location,
              skuName: wizardContext.iotHubDetails.sku?.name,
              skuTier: wizardContext.iotHubDetails.sku?.tier,
              hostName: wizardContext.iotHubDetails.properties?.hostName,
            }
          : undefined,
      }),
      ...(wizardContext.eventhouseValidated && {
        eventhouse: {
          eventhouseId: wizardContext.eventhouseId,
          eventhouseName: wizardContext.eventhouseName,
          databaseId: wizardContext.databaseId,
          databaseName: wizardContext.databaseName,
          queryServiceUri: wizardContext.queryServiceUri,
          ingestionServiceUri: wizardContext.ingestionServiceUri,
        },
      }),
      ...(wizardContext.tablesCreated && {
        tables: {
          telemetryTableName: wizardContext.telemetryTableName,
          propertiesTableName: wizardContext.propertiesTableName,
        },
      }),
      ...(wizardContext.eventstreamsCreated && {
        eventstreams: {
          telemetryStreamId: wizardContext.telemetryStreamId,
          telemetryStreamName: wizardContext.telemetryStreamName,
          propertiesStreamId: wizardContext.propertiesStreamId,
          propertiesStreamName: wizardContext.propertiesStreamName,
        },
      }),
      ...(wizardContext.routingConfigured && {
        routing: {
          telemetryRouteName: wizardContext.telemetryRouteName,
          propertiesRouteName: wizardContext.propertiesRouteName,
          telemetryEndpointName: wizardContext.telemetryEndpointName,
          propertiesEndpointName: wizardContext.propertiesEndpointName,
        },
      }),
      // Dashboard wizard fields
      ...(wizardContext.dashDashboardCreated && {
        dashboard: {
          dashboardId: wizardContext.dashDashboardId || "",
          dashboardName: wizardContext.dashDashboardName || "",
        },
      }),
      // Model wizard fields
      ...(wizardContext.modelUploaded && {
        deviceModel: {
          fileName: wizardContext.modelFileName || "",
          modelName: wizardContext.modelName || "",
          modelJson: wizardContext.modelJson || null,
          propsNormalizedName: wizardContext.propsNormalizedName || "",
          propsLkvViewName: wizardContext.propsLkvViewName || "",
          modeledDataName: wizardContext.modeledDataName || "",
        },
      }),
      // Activator wizard fields
      ...(wizardContext.activatorCreated && {
        activator: {
          activatorId: wizardContext.actActivatorId || "",
          activatorName: wizardContext.actActivatorName || "",
          eventstreamId: wizardContext.actEventstreamId || "",
          eventstreamName: wizardContext.actEventstreamName || "",
          connectionId: wizardContext.actConnectionId || "",
        },
      }),
      // Data Agent wizard fields
      ...(wizardContext.agentCreated && {
        dataAgent: {
          agentId: wizardContext.agentId || "",
          agentName: wizardContext.agentName || "",
          databaseId: wizardContext.agentDbId || "",
          databaseName: wizardContext.agentDbName || "",
          tableName: wizardContext.agentTableName || "",
        },
      }),
    };
    setCurrentDefinition(updatedDef);
    setSaveStatus(SaveStatus.NotSaved);

    // Notification labels
    const notifTitle = isDashboardWizard
      ? "Dashboard Created"
      : isModelWizard
        ? "Model Uploaded"
        : isActivatorWizard
          ? "Activator Created"
          : isDataAgentWizard
            ? "Data Agent Created"
            : "Setup Complete";
    const notifMsg = isDashboardWizard
      ? "Real-Time Dashboard created and saved."
      : isModelWizard
        ? "Device model uploaded and saved."
        : isActivatorWizard
          ? "Activator and Eventstream created and saved."
          : isDataAgentWizard
            ? "Data Agent created and configured with AI instructions."
            : "IoT Solution setup completed and saved.";

    // Auto-save after wizard completion
    try {
      await saveWorkloadItem<IoTSolutionItemDefinition>(workloadClient, {
        ...item,
        definition: updatedDef,
      });
      if (item) {
        item.definition = updatedDef;
      }
      setSaveStatus(SaveStatus.Saved);
      callNotificationOpen(workloadClient, notifTitle, notifMsg, undefined, undefined);
      setIsSavingWizard(false);
      // Navigate back to home cards page or workspace
      if (navigateHome) {
        navigateHome();
      } else if (item?.workspaceId) {
        navigateToWorkspace(workloadClient, item.workspaceId);
      }
    } catch (error) {
      setSaveStatus(SaveStatus.NotSaved);
      setIsSavingWizard(false);
      throw error;
    }
  };

  const isSaveEnabled = () => {
    return saveStatus !== SaveStatus.Saved && saveStatus !== SaveStatus.Saving;
  };

  // Wrapper that uses navigation hook for the home page
  const HomePageWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    return (
      <IoTSolutionItemHomePage
        workloadClient={workloadClient}
        item={item}
        onSelectWizard={(wizardId) => {
          if (wizardId === "ingestion") {
            setCurrentView(EDITOR_VIEW_TYPES.INGESTION_WIZARD);
          } else if (wizardId === "dashboard") {
            setCurrentView(EDITOR_VIEW_TYPES.DASHBOARD_WIZARD);
          } else if (wizardId === "model") {
            setCurrentView(EDITOR_VIEW_TYPES.MODEL_WIZARD);
          } else if (wizardId === "activator") {
            setCurrentView(EDITOR_VIEW_TYPES.ACTIVATOR_WIZARD);
          } else if (wizardId === "data-agent") {
            setCurrentView(EDITOR_VIEW_TYPES.DATA_AGENT_WIZARD);
          }
        }}
      />
    );
  };

  // Wrapper for ingestion wizard with back navigation
  const IngestionWizardWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    const goHome = () => setCurrentView(EDITOR_VIEW_TYPES.HOME);
    return (
      <IoTSolutionItemDefaultView
        workloadClient={workloadClient}
        item={item}
        onWizardComplete={(ctx) => handleWizardComplete(ctx, goHome)}
        onBack={goHome}
      />
    );
  };

  // Wrapper for dashboard wizard with back navigation
  const DashboardWizardWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    const goHome = () => setCurrentView(EDITOR_VIEW_TYPES.HOME);
    return (
      <IoTSolutionItemDashboardWizard
        workloadClient={workloadClient}
        item={item}
        onWizardComplete={(ctx) => handleWizardComplete(ctx, goHome)}
        onBack={goHome}
      />
    );
  };

  // Wrapper for model wizard with back navigation
  const ModelWizardWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    const goHome = () => setCurrentView(EDITOR_VIEW_TYPES.HOME);
    return (
      <IoTSolutionItemModelWizard
        workloadClient={workloadClient}
        item={item}
        onWizardComplete={(ctx) => handleWizardComplete(ctx, goHome)}
        onBack={goHome}
      />
    );
  };

  // Wrapper for activator wizard with back navigation
  const ActivatorWizardWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    const goHome = () => setCurrentView(EDITOR_VIEW_TYPES.HOME);
    return (
      <IoTSolutionItemActivatorWizard
        workloadClient={workloadClient}
        item={item}
        onWizardComplete={(ctx) => handleWizardComplete(ctx, goHome)}
        onBack={goHome}
      />
    );
  };

  // Wrapper for Data Agent wizard with back navigation
  const DataAgentWizardWrapper = () => {
    const { setCurrentView } = useViewNavigation();
    const goHome = () => setCurrentView(EDITOR_VIEW_TYPES.HOME);
    return (
      <IoTSolutionItemDataAgentWizard
        workloadClient={workloadClient}
        item={item}
        onWizardComplete={(ctx) => handleWizardComplete(ctx, goHome)}
        onBack={goHome}
      />
    );
  };

  const views = [
    {
      name: EDITOR_VIEW_TYPES.HOME,
      component: <HomePageWrapper />,
    },
    {
      name: EDITOR_VIEW_TYPES.INGESTION_WIZARD,
      component: <IngestionWizardWrapper />,
    },
    {
      name: EDITOR_VIEW_TYPES.DASHBOARD_WIZARD,
      component: <DashboardWizardWrapper />,
    },
    {
      name: EDITOR_VIEW_TYPES.MODEL_WIZARD,
      component: <ModelWizardWrapper />,
    },
    {
      name: EDITOR_VIEW_TYPES.ACTIVATOR_WIZARD,
      component: <ActivatorWizardWrapper />,
    },
    {
      name: EDITOR_VIEW_TYPES.DATA_AGENT_WIZARD,
      component: <DataAgentWizardWrapper />,
    },
  ];

  useEffect(() => {
    if (!isLoading && item && viewSetter) {
      viewSetter(EDITOR_VIEW_TYPES.HOME);
    }
  }, [isLoading, item, viewSetter]);

  return (
    <>
      {isSavingWizard && (
        <div style={{
          position: "fixed",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          zIndex: 99999,
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          gap: "16px",
          backgroundColor: "rgba(255, 255, 255, 0.97)",
        }}>
          <Spinner size="huge" label="Saving your configuration..." labelPosition="below" />
        </div>
      )}
      <ItemEditor
      isLoading={isLoading}
      loadingMessage="Loading IoT Solution..."
      ribbon={(context) => (
        <IoTSolutionItemRibbon
          {...props}
          viewContext={context}
          isSaveButtonEnabled={isSaveEnabled()}
          saveItemCallback={saveItem}
          openSettingsCallback={handleOpenSettings}
        />
      )}
      messageBar={[]}
      views={views}
      viewSetter={(setCurrentView) => {
        if (!viewSetter) {
          setViewSetter(() => setCurrentView);
        }
      }}
    />
    </>
  );
}
