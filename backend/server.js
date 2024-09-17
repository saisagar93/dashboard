const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const axios = require('axios');
const cors = require('cors');
const path = require('path');
const app = express();
const port = 3000;

// Serve static files from the React app
app.use(express.static(path.join(__dirname, 'build')));

// Middleware
app.use(cors());


// MSAL configuration with hard-coded credentials
const config = {
    auth: {
        clientId: '301bb42a-89af-4d2b-9d02-145150e7a776', // Replace with your actual client ID
        authority: 'https://login.microsoftonline.com/a4a817ea-2176-4b17-b6de-e09407c266e3', // Replace with your actual tenant ID
        clientSecret: '_eM8Q~3a8_RbB3~trxxnMLTzy6P-YTIjFBCQEcxX', // Replace with your actual client secret
    }
};

const cca = new ConfidentialClientApplication(config);

// Function to get access token
async function getAccessToken() {
    const tokenRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
    };

    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return response.accessToken;
    } catch (error) {
        console.error('Error acquiring token:', error);
        throw new Error('Error acquiring token');
    }
}
// Function to get all list items by list ID
async function getAllListItems(siteId, listId, allItems = []) {
    const accessToken = await getAccessToken();
    let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields`;

    try {
        let response;
        do {
            response = await axios.get(url, {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    Accept: 'application/json',
                }
            });

            const itemsWithFields = response.data.value.map(item => {
                const fields = item.fields;
                const fieldsToCheck = ['field_3', 'field_8', 'field_10'];
                console.log('date',fields);

                // No formatting needed; just ensure the fields are directly used as they are
                Object.keys(fields).forEach(key => {
                    if (fieldsToCheck.includes(key) && fields[key] && fields[key].match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/)) {
                        // fields[key] is kept as UTC date string
                        // No formatting is done here
                    }
                });

                return fields;
            });
            allItems = allItems.concat(itemsWithFields);

            url = response.data['@odata.nextLink'] || null;
        } while (url);

        return {
            items: allItems
        };
    } catch (error) {
        console.error('Error fetching list items:', error.response ? error.response.data : error.message);
        throw new Error('Error fetching list items');
    }
}

// SSE Endpoint
app.get('/events', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    const siteId = 'almadinalogistics.sharepoint.com,c3aa28b9-cc9f-4bb3-b9f4-623d3260b6e4,ac7e2e52-dcbc-4ffa-8bf5-ce2953c2c744'; // Replace with your SharePoint site ID
    const listId = 'ca3cbd16-e08c-4558-b113-a51c92ed57df'; // Replace with your list ID

    const sendUpdate = async () => {
        try {
            const data = await getAllListItems(siteId, listId);
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        } catch (error) {
            console.error('Error sending SSE update:', error);
        }
    };


    const intervalId = setInterval(sendUpdate, 200000);

    // Initial data send
    sendUpdate();

    // Clean up when client disconnects
    req.on('close', () => {
        clearInterval(intervalId);
        res.end();
    });
});

app.get('/list-items/:listId', async (req, res) => {
    const siteId = 'almadinalogistics.sharepoint.com,c3aa28b9-cc9f-4bb3-b9f4-623d3260b6e4,ac7e2e52-dcbc-4ffa-8bf5-ce2953c2c744'; // Replace with your SharePoint site ID
    const listId = req.params.listId;

    try {
        const response = await getAllListItems(siteId, listId);
        res.json(response);
    } catch (error) {
        console.error('Error retrieving list items:', error.message);
        res.status(500).send('Error retrieving list items');
    }
});

app.listen(port, () => console.log(`Server running on port ${port}`));
