# Initial Email Outlook Add-in

An Outlook Add-in that allows you to extract the body text of a selected email, save it as a JSON file, and send it to Azure OpenAI for analysis.

## Features

- Extract email body content from the selected email in Outlook
- Save email data (subject, sender, date, body) as a JSON file
- Send email data to Azure OpenAI for analysis and summarization
- Works with Outlook Desktop (Windows/Mac) and Outlook Web

## Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- Microsoft Outlook (Desktop or Web)
- OpenSSL (for certificate generation)
- Azure OpenAI API access

## Setup for Development

1. Clone or download this repository
2. Navigate to the project directory
3. Install the dependencies:
   ```
   npm install
   ```
4. Generate self-signed certificates for HTTPS development:
   ```
   # Create certs directory
   mkdir -p certs
   cd certs

   # Generate server key and certificate
   openssl req -newkey rsa:2048 -nodes -keyout server.key -x509 -days 365 -out server.crt -subj "/CN=localhost"

   # Create certificate serial number file
   echo "01" > ca.srl

   # Generate CA key
   openssl genrsa -out ca.key 2048

   # Generate CA certificate
   openssl req -new -x509 -key ca.key -out ca.crt -days 365 -subj "/CN=localhost"

   # Return to project root
   cd ..
   ```
5. Configure Azure OpenAI credentials:
   - Create a `.env` file in the project root
   - Add your Azure OpenAI credentials:
     ```
     AZURE_OPENAI_ENDPOINT=your_endpoint_here  # e.g. https://your-resource-name.openai.azure.com/
     AZURE_OPENAI_API_KEY=your_api_key_here
     AZURE_OPENAI_DEPLOYMENT_NAME=your_deployment_name_here
     ```
6. Start the development server:
   ```
   npm start
   ```

## Azure OpenAI Setup

1. Create an Azure OpenAI resource in the [Azure Portal](https://portal.azure.com)
2. Deploy a model in your Azure OpenAI resource
3. Get your endpoint URL and API key from the Azure Portal
4. Add these credentials to your `.env` file as shown above

## Sideloading the Add-in in Outlook

### Outlook Desktop (Windows)
1. Open Outlook
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

### Outlook Desktop (Mac)
1. Open Outlook
2. Click "Tools" > "Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

### Outlook Web
1. Go to [Outlook Web](https://outlook.office.com)
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

## Usage

1. Open Outlook and select an email
2. Click the "Initial Email" button in the ribbon
3. In the taskpane that appears:
   - Click "Save Email as JSON" to download the email content as a JSON file
   - Click "Send to Azure OpenAI" to analyze the email using Azure OpenAI

## Troubleshooting

### Certificate Issues
If you encounter certificate-related errors:
1. Make sure you've generated all the required certificate files in the `certs` directory
2. Verify that the certificates are valid and not expired
3. Check that the certificates are properly referenced in your development environment

### Azure OpenAI Issues
If you encounter issues with the Azure OpenAI integration:
1. Verify that your Azure OpenAI credentials in the `.env` file are correct
2. Check that your Azure OpenAI deployment is active and available
3. Ensure that you have sufficient quota available for your Azure OpenAI resource
4. Check the browser console for any API-related error messages

## License

ISC
