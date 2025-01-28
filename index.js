const https = require('https');
const getAccessToken = require('./auth');
require('dotenv').config();

async function performSearch(query) {
    try {
        // Get the access token
        const accessToken = await getAccessToken();

        // Define the request body for the search
        const requestBody = JSON.stringify({
            requests: [
                {
                    entityTypes: ['message', 'driveItem', 'chatMessage', 'user'], // Entities to search
                    query: {
                        queryString: query, // Search query
                    },
                    from: 0, // Pagination: start index
                    size: 25, // Number of results to return
                },
            ],
        });

        // Define the API endpoint
        const options = {
            hostname: 'graph.microsoft.com',
            path: '/v1.0/search/query',
            method: 'POST',
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
                'Content-Length': requestBody.length,
            },
        };

        // Make the HTTPS request
        const request = https.request(options, (response) => {
            let data = '';

            // Collect data chunks
            response.on('data', (chunk) => {
                data += chunk;
            });

            // Process the complete response
            response.on('end', () => {
                if (response.statusCode === 200) {
                    const searchResults = JSON.parse(data);
                    const formattedResults = formatSearchResults(searchResults);
                    console.log('Formatted Search Results:', formattedResults);
                } else {
                    console.error('Error performing search:', response.statusCode, data);
                }
            });
        });

        // Handle request errors
        request.on('error', (error) => {
            console.error('Request error:', error);
        });

        // Send the request body
        request.write(requestBody);
        request.end();
    } catch (error) {
        console.error('Error:', error);
    }
}

// Function to format search results
function formatSearchResults(searchResults) {
    const formattedResults = {
        Emails: [],
        OneDrive: [],
        Teams: [],
        Users: [],
    };

    // Process each search result
    searchResults.value[0].hitsContainers[0].hits.forEach((hit) => {
        const resource = hit.resource;

        switch (hit.resource['@odata.type']) {
            case '#microsoft.graph.message': // Email
                formattedResults.Emails.push({
                    Subject: resource.subject,
                    Date: resource.receivedDateTime,
                });
                break;

            case '#microsoft.graph.driveItem': // OneDrive file
                formattedResults.OneDrive.push({
                    Folder: resource.parentReference.path,
                    Name: resource.name,
                });
                break;

            case '#microsoft.graph.chatMessage': // Teams message
                formattedResults.Teams.push({
                    Channel: resource.channelIdentity.channelName,
                    From: resource.from.user.displayName,
                    Date: resource.createdDateTime,
                });
                break;

            case '#microsoft.graph.user': // User
                formattedResults.Users.push({
                    "Display Name": resource.displayName,
                    "Email": resource.mail || resource.userPrincipalName,
                });
                break;
        }
    });

    return formattedResults;
}

// Run the search function
const searchQuery = 'test'; // Replace with your search query
performSearch(searchQuery);