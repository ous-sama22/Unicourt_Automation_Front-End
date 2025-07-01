# UniCourt Automation Frontend - Google Apps Script

A Google Apps Script-based frontend interface for the UniCourt Automation Backend API, providing seamless integration with Google Sheets for case processing, data management, and automated document analysis.

## 🔗 Related Repositories

- **This Repository**: [Unicourt_Automation_Front-End](https://github.com/ous-sama22/Unicourt_Automation_Front-End) - Google Apps Script frontend for Google Sheets integration
- **Backend Repository**: [UniCourtAutomationBack-end---V4](https://github.com/ous-sama22/UniCourtAutomationBack-end---V4) - FastAPI backend service

## 🌐 Quick Links

- **📖 User Guide**: Open Google Sheets → Extensions → Apps Script → Open `Unicourt Processor UserGuide.html`
- **⚙️ Configuration**: Google Sheets → UniCourt Processor → Configure Settings
- **📊 Status Dashboard**: Google Sheets → UniCourt Processor → View Submission Log
- **🔧 Backend API Docs**: Your backend URL + `/docs` (e.g., `http://localhost:8000/docs`)

## 📋 Table of Contents

- [🌐 Quick Links](#-quick-links)
- [✨ Features](#-features)
- [🏗️ Architecture](#️-architecture)
- [🚀 Installation](#-installation)
- [⚙️ Configuration](#️-configuration)
- [📖 Usage](#-usage)
- [📁 File Structure](#-file-structure)
- [🔄 Workflow](#-workflow)
- [🐛 Troubleshooting](#-troubleshooting)
- [🛠️ Development](#️-development)
- [🔒 Security](#-security)
- [🤝 Contributing](#-contributing)

## ✨ Features

### Core Functionality
- **📊 Google Sheets Integration**: Native integration with Google Sheets for case data management
- **🤖 Automated Case Processing**: Submit cases to backend API for automated UniCourt processing
- **📈 Real-time Status Tracking**: Monitor case processing status and progress
- **📋 Batch Operations**: Process multiple cases simultaneously with batch operations
- **📝 Comprehensive Logging**: Track all submission attempts and API interactions

### User Interface Features
- **🎛️ Settings Management**: Easy configuration of backend connection and credentials
- **📊 Interactive Dashboard**: View case statuses, processing results, and logs
- **🔄 Auto-refresh**: Automatic status updates and data synchronization
- **🎨 Material Design UI**: Modern, responsive sidebar interface
- **📱 Mobile-friendly**: Works on mobile devices through Google Sheets app

### Data Management
- **🗂️ Sheet Structure Management**: Automatic sheet creation and header management
- **📋 Case Import/Export**: Easy data import from various sources
- **🔍 Search and Filter**: Built-in search and filtering capabilities
- **📊 Progress Tracking**: Visual progress indicators and status updates
- **📈 Reporting**: Generate reports on case processing metrics

## 🏗️ Architecture

### Component Overview
```
Google Apps Script Project
├── Unicourt Processor Main Code.gs     # Core application logic
├── Unicourt Processor BackendService.gs # API communication layer
├── Unicourt Processor Sheetmanager.gs  # Sheet operations and data management
├── Unicourt Processor DialogDisplay.html # Dialog interfaces
├── Unicourt Processor UISidebar.html   # Main settings sidebar
├── Unicourt Processor UserGuide.html   # Comprehensive user documentation
├── Unicourt Processor AutoSubmitIntervalDialog.html # Auto-submission settings
├── onOpen.gs                           # Google Sheets menu integration
└── read_files.py                       # Development utility
```

### Integration Flow
```
Google Sheets ←→ Apps Script ←→ Backend API ←→ UniCourt.com
     ↓              ↓              ↓              ↓
  User Data    Processing     Automation    Document
  Management    Logic        & Analysis    Extraction
```

### Key Components

#### 1. **Main Code (`Unicourt Processor Main Code.gs`)**
- Core application logic and workflow management
- User interface event handlers
- Settings management and persistence
- Error handling and logging

#### 2. **Backend Service (`Unicourt Processor BackendService.gs`)**
- API communication with the backend
- HTTP request handling and authentication
- Error logging and retry mechanisms
- Response processing and validation

#### 3. **Sheet Manager (`Unicourt Processor Sheetmanager.gs`)**
- Google Sheets operations and data management
- Sheet structure creation and maintenance
- Data validation and formatting
- Batch operations and updates

#### 4. **User Interface Components**
- **Settings Sidebar**: Configuration management interface
- **Dialog System**: Interactive dialogs for specific operations
- **User Guide**: Comprehensive documentation and help system

## 🚀 Installation

### Prerequisites
- Google Account with access to Google Sheets
- Running instance of the UniCourt Automation Backend API
- UniCourt.com account credentials
- OpenRouter API key for LLM processing

### Installation Steps

#### 1. **Create Google Sheets Document**
```
1. Go to Google Sheets (sheets.google.com)
2. Create a new spreadsheet
3. Name it "UniCourt Case Processing" (or your preferred name)
```

#### 2. **Set Up Apps Script Project**
```
1. In your Google Sheet: Extensions → Apps Script
2. Delete the default Code.gs file
3. Create new files for each component (see File Structure section)
4. Copy the code from each corresponding .gs and .html file
5. Save the project with a descriptive name
```

#### 3. **Configure Permissions**
```
1. In Apps Script Editor: Click "Run" on any function
2. Grant necessary permissions when prompted:
   - Google Sheets access
   - External URL access (for backend API calls)
   - Script properties access
```

#### 4. **Initial Setup**
```
1. Return to your Google Sheet
2. You should see a new "UniCourt Processor" menu
3. Click: UniCourt Processor → Ensure/Reset Sheet Headers & Structure
4. This creates all necessary sheets and headers
```

## ⚙️ Configuration

### Backend Connection Settings

#### 1. **Access Configuration**
```
1. In Google Sheets: UniCourt Processor → Configure Settings
2. In the sidebar, configure Backend Connection:
   - Backend API URL: http://your-server-ip:8000/api/v1
   - Backend API Key: Your backend API access key
3. Click "Save Backend Connection"
```

#### 2. **Client Credentials & Settings**
Configure the following in the Settings sidebar:
- **UniCourt Email**: Your UniCourt login email
- **UniCourt Password**: Your UniCourt login password
- **OpenRouter API Key**: Your OpenRouter API key
- **OpenRouter LLM Model**: Model name (e.g., `google/gemini-2.0-flash-001`)
- **Extract Associated Party Addresses**: Enable/disable address extraction

### Auto-Submission Settings
```
1. UniCourt Processor → Configure Auto-Submit Interval
2. Set interval in minutes (minimum 30 minutes)
3. Configure submission limits per batch
```

### Sheet Structure

The application automatically creates these sheets:
- **Cases**: Main case data and status tracking
- **Final Judgments**: Extracted final judgment information
- **Submission Log**: History of all submission attempts
- **Error Log**: API errors and troubleshooting information

## 📖 Usage

### Basic Workflow

#### 1. **Prepare Case Data**
In the "Cases" sheet, add your case information:
- **case_number_for_db_id**: Unique identifier for the case
- **case_name_for_search**: Name to search for on UniCourt
- **input_creditor_name**: Creditor name for matching
- **is_business**: TRUE/FALSE for business vs individual
- **creditor_type**: "business" or "individual"

#### 2. **Submit Cases for Processing**
```
1. Select cases to process (check the checkboxes)
2. UniCourt Processor → Submit Selected Cases
3. Monitor progress in the status columns
```

#### 3. **Monitor Progress**
- **Status Columns**: Real-time status updates
- **Submission Log**: View all submission attempts
- **Error Log**: Check for any processing issues

#### 4. **View Results**
- **Cases Sheet**: Updated with processing results
- **Final Judgments Sheet**: Extracted judgment information
- **Downloaded Documents**: Available through backend file system

### Advanced Features

#### **Batch Operations**
- **Refresh All Statuses**: Update status for all cases
- **Submit All Pending**: Process all unprocessed cases
- **Clear Selections**: Reset all checkboxes

#### **Auto-Submission**
- Configure automatic submission intervals
- Set processing limits per batch
- Monitor auto-submission logs

#### **Configuration Management**
- **Show Current Backend Config**: View backend settings
- **Request Backend Restart**: Restart backend service
- **Save All Settings**: Comprehensive configuration save

## 📁 File Structure

### Core Files

| File | Purpose | Key Functions |
|------|---------|---------------|
| `Unicourt Processor Main Code.gs` | Core logic and settings | `submitSelectedCases()`, `refreshAllStatuses()`, configuration management |
| `Unicourt Processor BackendService.gs` | API communication | `callBackendApi()`, error logging, request handling |
| `Unicourt Processor Sheetmanager.gs` | Sheet operations | Sheet creation, data management, formatting |
| `onOpen.gs` | Menu integration | Creates UniCourt Processor menu in Google Sheets |

### User Interface Files

| File | Purpose | Description |
|------|---------|-------------|
| `Unicourt Processor UISidebar.html` | Main settings interface | Configuration sidebar with Material Design |
| `Unicourt Processor DialogDisplay.html` | Dialog system | Modal dialogs for specific operations |
| `Unicourt Processor UserGuide.html` | Documentation | Comprehensive user guide and help |
| `Unicourt Processor AutoSubmitIntervalDialog.html` | Auto-submission settings | Configure automatic processing intervals |

### Configuration Structure

#### Apps Script Properties
The application stores configuration in Apps Script Properties:
- `BACKEND_URL`: Backend API base URL
- `BACKEND_API_KEY`: API authentication key
- `UNICOURT_EMAIL`: UniCourt login email
- `UNICOURT_PASSWORD`: UniCourt password (encrypted)
- `OPENROUTER_KEY`: OpenRouter API key (encrypted)
- `OPENROUTER_MODEL`: LLM model selection
- `EXTRACT_ASSOCIATED_PARTY_ADDRESSES`: Address extraction setting

## 🔄 Workflow

### Case Processing Flow
```
1. User Input → Cases Sheet
2. Case Selection → Submit to Backend
3. Backend Processing → UniCourt Automation
4. Document Analysis → LLM Processing
5. Results Storage → Sheet Updates
6. Status Tracking → Real-time Updates
```

### Error Handling Flow
```
1. API Error Detection → Error Logging
2. Retry Mechanisms → Automatic Retries
3. User Notification → Status Updates
4. Troubleshooting → Error Log Review
```

## 🐛 Troubleshooting

### Common Issues

#### **Backend Connection Issues**
- **Error**: "Backend URL or API Key not configured"
- **Solution**: Configure backend connection in Settings sidebar
- **Check**: Verify backend is running and accessible

#### **API Authentication Errors**
- **Error**: "Invalid or missing API Key"
- **Solution**: Verify API key in backend connection settings
- **Check**: Ensure API key matches backend configuration

#### **Case Submission Failures**
- **Error**: Cases stuck in "pending" status
- **Solution**: Check Error Log for specific issues
- **Check**: Verify UniCourt credentials and backend health

#### **Sheet Structure Issues**
- **Error**: Missing columns or sheets
- **Solution**: Run "Ensure/Reset Sheet Headers & Structure"
- **Check**: Verify all required sheets exist

### Debugging Tools

#### **Error Log Viewing**
```
1. UniCourt Processor → View Submission Log
2. Check the Error Log section
3. Review timestamps and error details
```

#### **Backend Health Check**
```
1. UniCourt Processor → Configure Settings
2. Click "Show Current Backend Config"
3. Verify backend connectivity and status
```

#### **Manual Status Refresh**
```
1. Select specific cases
2. UniCourt Processor → Refresh Selected Statuses
3. Check for updated information
```

### Performance Optimization

#### **Batch Size Optimization**
- Adjust auto-submission batch sizes based on backend capacity
- Monitor processing times and adjust intervals accordingly
- Use smaller batches for complex cases

#### **Network Optimization**
- Ensure stable internet connection
- Consider backend server location for latency
- Monitor API response times

## 🛠️ Development

### Development Setup

#### **Local Development**
```
1. Clone the repository
2. Set up your development backend instance
3. Create a test Google Sheet
4. Deploy code changes through Apps Script Editor
```

#### **Code Structure**
- Follow Google Apps Script best practices
- Use clear function naming conventions
- Implement comprehensive error handling
- Document all public functions

#### **Testing**
- Test with small case batches first
- Verify all sheet operations
- Test error scenarios
- Validate API integrations

### Customization

#### **Adding New Features**
1. **New API Endpoints**: Update `BackendService.gs`
2. **Sheet Operations**: Modify `Sheetmanager.gs`
3. **UI Components**: Update HTML files
4. **Core Logic**: Extend `Main Code.gs`

#### **Configuration Options**
- Add new settings to `UNICOURT_CONFIG`
- Update settings UI in sidebar
- Implement backend configuration sync

## 🔒 Security

### Data Protection
- **Credentials**: Stored securely in Apps Script Properties
- **API Keys**: Never exposed in client-side code
- **Passwords**: Stored encrypted and never returned to UI
- **Transmission**: All API calls use HTTPS

### Access Control
- **Google Account**: Requires Google account authentication
- **Sheet Permissions**: Controlled through Google Sheets sharing
- **API Access**: Secured through backend API key authentication

### Best Practices
- **Regular Updates**: Keep credentials updated and secure
- **Access Review**: Regularly review who has access to sheets
- **API Monitoring**: Monitor API usage and errors
- **Backup**: Regular backup of important case data

## 🤝 Contributing

### Development Guidelines

1. **Code Style**
   - Follow Google Apps Script conventions
   - Use clear, descriptive function names
   - Comment complex logic thoroughly
   - Maintain consistent indentation

2. **Testing**
   - Test all changes with sample data
   - Verify error handling scenarios
   - Test UI interactions thoroughly
   - Validate API integration points

3. **Documentation**
   - Update README for new features
   - Document configuration changes
   - Update user guide for UI changes
   - Include troubleshooting information

### Contribution Process

1. **Fork the Repository**
   ```
   git clone https://github.com/ous-sama22/Unicourt_Automation_Front-End.git
   ```

2. **Create Feature Branch**
   ```
   git checkout -b feature/amazing-feature
   ```

3. **Make Changes**
   - Implement your feature or fix
   - Test thoroughly with sample data
   - Update documentation as needed

4. **Submit Pull Request**
   ```
   git commit -m 'Add amazing feature'
   git push origin feature/amazing-feature
   ```

### Feature Requests

1. Check [GitHub Issues](https://github.com/ous-sama22/Unicourt_Automation_Front-End/issues) for existing requests
2. Create detailed issue with use case description
3. Include mockups or examples if applicable
4. Consider implementation complexity and maintenance

## 📄 License

This project is licensed under the MIT License. See the LICENSE file for details.