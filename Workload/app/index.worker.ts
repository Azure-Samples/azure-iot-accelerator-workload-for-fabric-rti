// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
    createWorkloadClient,
    InitParams,
    ItemLikeV2,
    NotificationType
} from '@ms-fabric/workload-client';
import { callPageOpen } from './controller/PageController';

interface ItemCreationFailureData {
    errorCode?: string;
    resultCode?: string;
}

interface ItemCreationSuccessData {
    item: ItemLikeV2;
}


export async function initialize(params: InitParams) {
    console.log('🚀 Worker initialization started with params:', params);

    const workloadClient = createWorkloadClient();
    console.log('✅ WorkloadClient created successfully');

    const workloadName = process.env.WORKLOAD_NAME;

    workloadClient.action.onAction(async function ({ action, data }) {
        console.log(`🧭 Started action ${action} with data:`, data);
        switch (action) {
            case 'item.onCreationSuccess':
                const { item: createdItem } = data as ItemCreationSuccessData;
                var path = "/item-editor";
                const itemTypeName = createdItem.itemType.substring(createdItem.itemType.lastIndexOf('.') + 1);
                path = `/${itemTypeName}Item-editor`;
                console.log(`Item created successfully, redirecting to ${path}/${createdItem.objectId}`);
                await callPageOpen(workloadClient, workloadName, `${path}/${createdItem.objectId}`);
                return Promise.resolve({ succeeded: true });

            case 'item.onCreationFailure': {
                const failureData = data as ItemCreationFailureData;
                await workloadClient.notification.open(
                    {
                        title: 'Error creating item',
                        notificationType: NotificationType.Error,
                        message: `Failed to create item, error code: ${failureData.errorCode}, result code: ${failureData.resultCode}`
                    });
                return Promise.resolve({ succeeded: false });
            }
            case 'getItemSettings': {
                return [];
            }
            default:
                throw new Error('Unknown action received');
        }
    });
}
