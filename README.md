# Microsoft Graph API Search Integration

This project demonstrates how to use the **Microsoft Graph API** to search across **Emails**, **OneDrive files**, **Teams messages**, and **Users**. The results are formatted into key-value pairs for easy readability.

## Features
- Search across multiple Microsoft 365 services:
  - **Emails**: Retrieve email subjects and dates.
  - **OneDrive**: Retrieve folder paths and file names.
  - **Teams**: Retrieve channel names, sender names, and message dates.
  - **Users**: Retrieve user display names and email addresses.
- Built using Node.js and the `https` module.
- Uses the Microsoft Authentication Library (`@azure/msal-node`) for authentication.

## Prerequisites
1. **Node.js**: Ensure Node.js is installed on your machine. Download it from [nodejs.org](https://nodejs.org/).
2. **Azure AD App Registration**:
   - Register an application in the [Azure Portal](https://portal.azure.com/).
   - Note down the **Client ID**, **Tenant ID**, and **Client Secret**.
   - Grant the following API permissions:
     - `Mail.Read` (for emails)
     - `Files.Read.All` (for OneDrive files)
     - `Chat.Read` (for Teams messages)
     - `User.Read.All` (for users)
   - Ensure admin consent is granted for the permissions.

## Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/microsoft-graph-search.git
   cd microsoft-graph-search
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Configure the application:
   - Create a `.env` file in the root directory and add your Azure AD credentials:
     ```env
     CLIENT_ID=your-client-id
     TENANT_ID=your-tenant-id
     CLIENT_SECRET=your-client-secret
     ```

4. Run the application:
   ```bash
   node index.js
   ```

## Usage
The application performs a search across Emails, OneDrive files, Teams messages, and Users. By default, it searches for the keyword `test`. You can modify the search query in the `index.js` file.

### Example Output
The search results are formatted into key-value pairs:
```json
{
  "Emails": [
    { "Subject": "Test Email", "Date": "2023-10-01T12:34:56Z" }
  ],
  "OneDrive": [
    { "Folder": "/drive/root:/Documents", "Name": "Report.pdf" }
  ],
  "Teams": [
    { "Channel": "General", "From": "John Doe", "Date": "2023-10-01T12:34:56Z" }
  ],
  "Users": [
    { "Display Name": "John Doe", "Email": "john.doe@example.com" }
  ]
}
```

## Code Structure
- **`auth.js`**: Handles authentication and retrieves an access token using `@azure/msal-node`.
- **`index.js`**: Performs the search and formats the results.
- **`.env`**: Stores sensitive configuration details (e.g., Client ID, Tenant ID, Client Secret).

## Dependencies
- [`@azure/msal-node`](https://www.npmjs.com/package/@azure/msal-node): Microsoft Authentication Library for Node.js.
- `https`: Node.js built-in module for making HTTP requests.

## Contributing
Contributions are welcome! If you find any issues or have suggestions for improvement, please open an issue or submit a pull request.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
