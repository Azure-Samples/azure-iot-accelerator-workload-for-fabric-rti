// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { Text } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import "./IoTSolutionItem.scss";

/* eslint-disable @typescript-eslint/no-var-requires */
const ingestionCardImg = require("../../assets/items/IoTSolutionItem/cards/ingestion-card.png");
const dashboardCardImg = require("../../assets/items/IoTSolutionItem/cards/generate-dashboard.png");
const modelCardImg = require("../../assets/items/IoTSolutionItem/cards/modeled-data.png");
const activatorCardImg = require("../../assets/items/IoTSolutionItem/cards/activator-rules.png");
const dataAgentCardImg = require("../../assets/items/IoTSolutionItem/cards/data-agent.png");

interface WizardCard {
  id: string;
  title: string;
  description: string;
  requirement: string;
  image: string;
  enabled: boolean;
}

interface WizardCardSection {
  id: string;
  title: string;
  description: string;
  cards: WizardCard[];
}

interface IoTSolutionItemHomePageProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onSelectWizard: (wizardId: string) => void;
}

/**
 * Home page for the IoT Solution item.
 * Groups setup and solution experiences by purpose and dependency.
 */
export function IoTSolutionItemHomePage({
  onSelectWizard,
}: IoTSolutionItemHomePageProps) {
  const sections: WizardCardSection[] = [
    {
      id: "foundation",
      title: "Set up your IoT data foundation",
      description:
        "Connect device data to Fabric first. Modeling is optional and can be added when you need typed fields, semantic context, and current device state.",
      cards: [
        {
          id: "ingestion",
          title: "Ingest device data from your IoT Hub",
          description:
            "Use Eventstreams to continuously ingest device telemetry and properties into an Eventhouse for real-time analytics.",
          requirement: "Start here",
          image: ingestionCardImg,
          enabled: true,
        },
        {
          id: "model",
          title: "Model your device data",
          description:
            "Optionally add semantic context and current device state to live telemetry for richer monitoring, alerts, and AI experiences.",
          requirement: "Optional enrichment",
          image: modelCardImg,
          enabled: true,
        },
      ],
    },
    {
      id: "experiences",
      title: "Create analytics, monitoring, and AI experiences",
      description:
        "Choose the experiences that support your goals. These options are independent and do not need to be configured in the order shown.",
      cards: [
        {
          id: "dashboard",
          title: "Generate Real-Time Dashboards",
          description:
            "Create dashboards that visualize live telemetry, recent values, and current device properties.",
          requirement: "Requires ingested data",
          image: dashboardCardImg,
          enabled: true,
        },
        {
          id: "activator",
          title: "Set up Activator rules",
          description:
            "Monitor modeled device data and current state, then trigger alerts, notifications, and automated response workflows.",
          requirement: "Requires modeled data",
          image: activatorCardImg,
          enabled: true,
        },
        {
          id: "data-agent",
          title: "Set up Data Agent",
          description:
            "Use AI to explore modeled real-time telemetry and current device state through natural-language questions.",
          requirement: "Requires modeled data",
          image: dataAgentCardImg,
          enabled: true,
        },
      ],
    },
  ];

  const renderCard = (card: WizardCard) => (
    <div
      key={card.id}
      className={`iot-solution-home-card ${!card.enabled ? "iot-solution-home-card--disabled" : ""}`}
      onClick={() => card.enabled && onSelectWizard(card.id)}
      onKeyDown={(e) => {
        if ((e.key === "Enter" || e.key === " ") && card.enabled) {
          e.preventDefault();
          onSelectWizard(card.id);
        }
      }}
      tabIndex={card.enabled ? 0 : -1}
      role="button"
      aria-label={card.title}
      aria-disabled={!card.enabled}
    >
      <div className="iot-solution-home-card-img-container">
        <img
          src={card.image}
          alt=""
          className="iot-solution-home-card-img"
        />
      </div>
      <div className="iot-solution-home-card-content">
        <span
          className={`iot-solution-home-card-requirement ${
            card.requirement.startsWith("Requires")
              ? "iot-solution-home-card-requirement--prerequisite"
              : ""
          }`}
        >
          {card.requirement}
        </span>
        <span className="iot-solution-home-card-title">{card.title}</span>
        <Text className="iot-solution-home-card-description">
          {card.description}
        </Text>
      </div>
    </div>
  );

  return (
    <div className="iot-solution-home">
      <div className="iot-solution-home-header">
        <h1 className="iot-solution-home-title">Azure IoT Solution Accelerator</h1>
        <Text className="iot-solution-home-description">
          Accelerate operational analytics, monitoring, alerting, and AI-powered
          insights using Azure IoT and Microsoft Fabric Real-Time Intelligence.
        </Text>
      </div>
      <div className="iot-solution-home-sections">
        {sections.map((section) => (
          <section
            key={section.id}
            className={`iot-solution-home-section iot-solution-home-section--${section.id}`}
          >
            <div className="iot-solution-home-section-header">
              <h2 className="iot-solution-home-section-title">{section.title}</h2>
              <Text className="iot-solution-home-section-description">
                {section.description}
              </Text>
            </div>
            <div className="iot-solution-home-cards">
              {section.cards.map(renderCard)}
            </div>
          </section>
        ))}
      </div>
    </div>
  );
}
