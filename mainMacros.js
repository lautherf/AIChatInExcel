/**
 * Define constants for AI agent and channel type.
 */

const AGENTTYPE_AIAgent = 'AIAgent'; // Constant for AI agent type
const CHANNEL_NORMAL = 'Channel'; // Constant for normal channel type


const AIAgentRegex = /AIAgent(.*?)/;
const ChannelRegex = /Channel(.*?)/;

// Declare variables at the top level of the script for use within any function
let curSheetType; // Placeholder for the current sheet type

/**
 * Reads configuration from a given range and returns a Map object.
 * @param {string} rangeText - The range from which to read the configuration.
 * @returns {Map} - A Map object containing the configuration data.
 */
function readConfig(rangeText) {
    // Create a new Range object from the provided text (e.g., 'Sheet1!A:B')
    let range = Application.Range(rangeText);

    // Create a new Map to store the configuration data.
    let rsMap = new Map();

    // Iterate over each row in the range.
    for (let rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++) {
        // Get the current row's Range object.
        let currentRow = range.Rows.Item(rowIndex);

        // Read the value from the first column of the current row.
        let firstColValue = currentRow.Cells.Item(1).Value();

        // Read the value from the second column of the current row.
        let secondColValue = currentRow.Cells.Item(2).Value();

        // If the first column's value is undefined, break the loop.
        if (undefined == firstColValue) {
            break;
        }

        // Add the key-value pair to the Map.
        rsMap.set(firstColValue, secondColValue);
    }

    // Return the Map containing the configuration data.
    return rsMap;
}

/**
 * Calls the chat API with the provided parameters.
 */
function callChatAPI() {
    // Get the configuration sheet name from the workspace configuration.
    var configSheetName = getWorkspaceConfig().get('Config_AI');

    // Read the configuration data from the specified range.
    let kimiConfig = readConfig(configSheetName + '!A:B');

    // Construct the HTTP request URL and parameters.
    var apiUrl = kimiConfig.get('apiUrl'); // API URL
    var apiKey = kimiConfig.get('apiKey'); // API key
    var model = kimiConfig.get('model'); // Model name
    var responseNum = kimiConfig.get('responseNum'); // Number of responses
    var temperature = kimiConfig.get('temperature'); // Temperature parameter

    var params = {
        "model": model,
        "n": responseNum,
        "messages": getConversationMessage(),
        "temperature": temperature
    };

    // Send the HTTP request to the chat API.
    kimi_XHR(apiUrl, apiKey, params);
}

/**
 * Sends an HTTP request to the specified API URL.
 * @param {string} apiUrl - The API URL to send the request to.
 * @param {string} apiKey - The API key for authentication.
 * @param {object} params - The parameters to send with the request.
 */
function kimi_XHR(apiUrl, apiKey, params) {
    // Print the request content to the console.
    Debug.Print('Request Content:' + JSON.stringify(params));

    // Create a new XMLHttpRequest object.
    var xhr = new XMLHttpRequest();
    xhr.open('POST', apiUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('Authorization', 'Bearer ' + apiKey);

    // Define a function to handle the response.
    xhr.onreadystatechange = function () {
        var readyState = xhr.readyState;
        var status = xhr.status;
        var responseText = xhr.responseText;
        var responseJson = JSON.parse(xhr.responseText);

        // If the request status is not 200, print the error message.
        if (xhr.status != 200) {
            var responseTextErrorMessage = responseJson.error.message;
            var responseTextErrorType = responseJson.error.type;
            Debug.Print('Exception:' + responseTextErrorMessage + responseTextErrorType);
            appendMessage('console', xhr.readyState + ' ' + responseTextErrorMessage + ' ' + responseTextErrorType);
        } else if (xhr.readyState === 4 && xhr.status === 200) {
            // If the request is complete and the status is 200, process the response.
            let choicesSize = responseJson.choices.length;
            deleteLastMessage();
            for (let i = 0; i < responseJson.choices.length; i++) {
                let responseMessageContent = responseJson.choices[i].message.content;
                Debug.Print('Response Text:' + xhr.responseText);
                appendMessage('assistant', responseMessageContent);
            }
            appendMessage('user', '');
        } else if (xhr.readyState === 2 || xhr.readyState === 3) {
            // If the request is processing, print a message to the console.
            Debug.Print('Processing...');
        }
    };

    // Send the request with the specified parameters.
    xhr.send(JSON.stringify(params));
    appendMessage('console', 'The request has been sent. If the content is extensive, it may take a few more seconds.');
}

// Define the function deleteLastMessage with no arguments
function deleteLastMessage() {
    // Get the name of the active sheet
    let ActiveSheetName = Application.Application.ActiveSheet.Name;

    // Define the range as columns A to C on the active sheet
    var rangeStr = ActiveSheetName + '!A:C';
    let range = Application.Range(rangeStr);

    // Initialize a variable to keep track of the last row with a value
    let endIndex = 0;

    // Iterate over each row in the range
    for (let rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++) {
        // Get the current row as a Range object
        let currentRow = range.Rows.Item(rowIndex);

        // Get the value of the cell in the first column of the current row
        let firstColValue = currentRow.Cells.Item(1).Value();
        // Get the value of the cell in the second column of the current row
        let secondColValue = currentRow.Cells.Item(2).Value();

        // If the first column is empty, we've reached the end of the data
        if (undefined == firstColValue) {
            break;
        } else {
            // Update the endIndex to the current row index
            endIndex = rowIndex;
        }
    }

    // Delete the row at the endIndex
    range.Rows.Item(endIndex).Delete();
}

// Appends a message to the end of a range in the active sheet
function appendMessage(role, text) {
    let ActiveSheetName = Application.Application.ActiveSheet.Name;
    var rangeStr = ActiveSheetName + '!A:B';
    let range = Application.Range(rangeStr);

    let endIndex = 0;

    // Iterate over each row in the range
    for (let rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++) {
        // Get the current row as a Range object
        let currentRow = range.Rows.Item(rowIndex);
        // Get the value in the first column of the current row
        let firstColValue = currentRow.Cells.Item(1).Value();

        if (undefined == firstColValue) {
            endIndex = rowIndex;
            break;
        }
    }

    // Insert the new message in the first empty row
    let firstEmptyRow = range.Rows.Item(endIndex);
    firstEmptyRow.Cells.Item(1, 1).Value2 = role;
    firstEmptyRow.Cells.Item(1, 2).Value2 = text;
}

// Returns the conversation messages from a Channel sheet
function getConversationMessage_Channel() {
    let ActiveSheetName = Application.Application.ActiveSheet.Name;
    let sheetMessageList = getSheetMessage(ActiveSheetName, 'A:C', 2);

    var lastItem = sheetMessageList[sheetMessageList.length - 1];
    let AIAgentName = lastItem['3'];
    let userMessageContent = lastItem['2'];

    let messageArray;

    if (undefined == AIAgentName) {
        messageArray = adapterAIAgentMessage(getAIAgentConversationMessageByName(ActiveSheetName));
    } else {
        messageArray = adapterAIAgentMessage(getAIAgentConversationMessageByName(AIAgentName));

        // If the AI agent has a default user line, remove it
        messageArray.pop();
        let obj = new Object();
        obj.role = 'user';
        obj.content = userMessageContent;
        messageArray.push(obj);
    }

    return messageArray;
}

// Returns the conversation messages from an AIAgent sheet
function getConversationMessage_AIAgent() {
    let ActiveSheetName = Application.Application.ActiveSheet.Name;
    return adapterAIAgentMessage(getAIAgentConversationMessageByName(ActiveSheetName));
}

// Formats the messages from the sheet into a consistent object format
function adapterAIAgentMessage(messageArray) {
    let rsArray = new Array();

    for (let i = 0; i < messageArray.length; i++) {
        let tempObj = messageArray[i];
        let tempObjRole = tempObj['1'];
        let tempObjContent = tempObj['2'];

        if ('system' == tempObjRole || 'user' == tempObjRole || 'assistant' == tempObjRole) {
            let obj = new Object();
            obj.role = tempObjRole;
            obj.content = tempObjContent;
            rsArray.push(obj);
        }
    }

    return rsArray;
}

// Retrieves messages from a specified range in a sheet
function getSheetMessage(sheetName, sheetColumn, startRowNum) {
    let rangeStr = sheetName + '!' + sheetColumn;
    let range = Application.Range(rangeStr);
    let messageArray = new Array();

    for (let rowIndex = startRowNum; rowIndex <= range.Rows.Count; rowIndex++) {
        let currentRow = range.Rows.Item(rowIndex);

        if (undefined == currentRow.Cells.Item(1).Value()) {
            break;
        }

        let obj = new Object();
        for (let columnIndex = 1; columnIndex <= range.Columns.Count; columnIndex++) {
            let columnIndexStr = '' + columnIndex;
            obj[columnIndexStr] = currentRow.Cells.Item(columnIndex).Value();
        }

        messageArray.push(obj);
    }

    return messageArray;
}

// Retrieves the conversation messages from an AIAgent sheet given the AIAgent's name
function getAIAgentConversationMessageByName(AIAgentName) {
    return getSheetMessage(AIAgentName, 'A:B', 2);
}

// Returns the conversation messages based on the type of the active sheet
function getConversationMessage() {
    let ActiveSheetName = Application.Application.ActiveSheet.Name;
    curSheetType = typeOfAgent(ActiveSheetName);

    if (AGENTTYPE_AIAgent == curSheetType) {
        return getConversationMessage_AIAgent();
    } else if (CHANNEL_NORMAL == curSheetType) {
        return getConversationMessage_Channel();
    }
}

// Determines the type of the active sheet (either an AIAgent sheet or a Channel sheet)


function typeOfAgent(name) {
    let rsStr;

    if (name.match(AIAgentRegex)) {
        rsStr = AGENTTYPE_AIAgent;
    } else if (name.match(ChannelRegex)) {
        rsStr = CHANNEL_NORMAL;
    }
    return rsStr;

}

function getWorkspaceConfig() {
    let workspaceConfig = readConfig('WorkspaceConfig!A:B');
    return workspaceConfig;
}

function CommandButtonClick() {
    callChatAPI()
}

function CommandButton1_Click() {
    CommandButtonClick();
}

function CommandButton2_Click() {
    CommandButtonClick();
}

function CommandButton3_Click() {
    CommandButtonClick();
}
/**
 * CommandButton4_Click Macro
 */
function CommandButton4_Click() {
    CommandButtonClick();
}
/**
 * CommandButton5_Click Macro
 */
function CommandButton5_Click() {
    CommandButtonClick();
}
/**
 * CommandButton6_Click Macro
 */
function CommandButton6_Click() {
    CommandButtonClick();
}

/**
 * CommandButton7_Click Macro
 */
function CommandButton7_Click() {
    CommandButtonClick();
}
/**
 * CommandButton8_Click Macro
 */
function CommandButton8_Click() {
    CommandButtonClick();

}
/**
 * CommandButton9_Click Macro
 */
function CommandButton9_Click() {
    CommandButtonClick();

}
/**
 * CommandButton10_Click Macro
 */
function CommandButton10_Click() {
    CommandButtonClick();
}
/**
 * CommandButton11_Click Macro
 */
function CommandButton11_Click() {
    CommandButtonClick();
}
/**
 * CommandButton12_Click Macro
 */
function CommandButton12_Click() {
    CommandButtonClick();
}
/**
 * CommandButton13_Click Macro
 */
function CommandButton13_Click() {
    CommandButtonClick();
}
/**
 * CommandButton14_Click Macro
 */
function CommandButton14_Click() {
    CommandButtonClick();
}
/**
 * CommandButton15_Click Macro
 */
function CommandButton15_Click() {
    CommandButtonClick();
}
/**
 * CommandButton16_Click Macro
 */
function CommandButton16_Click() {
    CommandButtonClick();
}
/**
 * CommandButton17_Click Macro
 */
function CommandButton17_Click() {
    CommandButtonClick();
}
/**
 * CommandButton18_Click Macro
 */
function CommandButton18_Click() {
    CommandButtonClick();
}
/**
 * CommandButton19_Click Macro
 */
function CommandButton19_Click() {
    CommandButtonClick();
}
/**
 * CommandButton20_Click Macro
 */
function CommandButton20_Click() {
    CommandButtonClick();
}
