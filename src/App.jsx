import React, { useState } from 'react';
import { Mail, CheckSquare, Calendar, AlertCircle, Play, Loader, ExternalLink } from 'lucide-react';

const OutlookTodoAgent = () => {
  const [todos, setTodos] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState('');
  const [accessToken, setAccessToken] = useState('');
  const [emailCount, setEmailCount] = useState(0);
  const [showSetup, setShowSetup] = useState(true);

  // Microsoft Graph API configuration
  const clientId = 'c401040d-0ed1-430e-abb4-91137450cdc8'; // Replace with your Azure AD app client ID
  const redirectUri = window.location.origin;
  const scopes = ['Mail.Read', 'User.Read'];

  // Authentication with Microsoft
  const authenticateWithMicrosoft = async () => {
    try {
      const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
        `client_id=${clientId}` +
        `&response_type=token` +
        `&redirect_uri=${encodeURIComponent(redirectUri)}` +
        `&scope=${encodeURIComponent(scopes.join(' '))}` +
        `&response_mode=fragment`;
      
      window.location.href = authUrl;
    } catch (err) {
      setError('Authentication failed: ' + err.message);
    }
  };

  // Check for access token in URL (after redirect)
  React.useEffect(() => {
    const hash = window.location.hash.substring(1);
    const params = new URLSearchParams(hash);
    const token = params.get('access_token');
    
    if (token) {
      setAccessToken(token);
      setShowSetup(false);
      window.history.replaceState({}, document.title, window.location.pathname);
    }
  }, []);

  // Extract to-do items from email content using Claude API
  const extractTodosFromEmail = async (emailSubject, emailBody) => {
    try {
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content: `Analyze this email and extract any action items, tasks, or to-dos. Return ONLY a JSON array of tasks, with no preamble or markdown. Each task should have: title (brief), description (optional details), priority (high/medium/low), and dueDate (if mentioned, in ISO format, otherwise null).

Email Subject: ${emailSubject}
Email Body: ${emailBody.substring(0, 2000)}

Return format: [{"title": "...", "description": "...", "priority": "medium", "dueDate": null}]`
            }
          ],
        })
      });

      const data = await response.json();
      const content = data.content[0].text;
      
      const cleanContent = content.replace(/```json|```/g, '').trim();
      const tasks = JSON.parse(cleanContent);
      
      return Array.isArray(tasks) ? tasks : [];
    } catch (err) {
      console.error('Error extracting todos:', err);
      return [];
    }
  };

  // Fetch emails from Outlook
  const fetchEmails = async () => {
    if (!accessToken) {
      setError('Please authenticate first');
      return;
    }

    setIsProcessing(true);
    setError('');
    setTodos([]);
    
    try {
      const twoMonthsAgo = new Date();
      twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
      const filterDate = twoMonthsAgo.toISOString();

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge ${filterDate}&$top=50&$select=subject,bodyPreview,body,receivedDateTime,from`,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (!response.ok) {
        throw new Error('Failed to fetch emails. Please re-authenticate.');
      }

      const data = await response.json();
      const emails = data.value;
      setEmailCount(emails.length);

      const allTodos = [];
      for (const email of emails) {
        const emailBody = email.body.content || email.bodyPreview;
        const extractedTodos = await extractTodosFromEmail(email.subject, emailBody);
        
        extractedTodos.forEach(todo => {
          allTodos.push({
            ...todo,
            emailSubject: email.subject,
            emailDate: new Date(email.receivedDateTime).toLocaleDateString(),
            from: email.from.emailAddress.name || email.from.emailAddress.address,
            id: Math.random().toString(36).substr(2, 9)
          });
        });
      }

      setTodos(allTodos);
    } catch (err) {
      setError(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const getPriorityColor = (priority) => {
    switch (priority?.toLowerCase()) {
      case 'high': return 'text-red-600 bg-red-50';
      case 'medium': return 'text-yellow-600 bg-yellow-50';
      case 'low': return 'text-green-600 bg-green-50';
      default: return 'text-gray-600 bg-gray-50';
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4 md:p-8">
      <div className="max-w-6xl mx-auto">
        <div className="bg-white rounded-2xl shadow-2xl overflow-hidden">
          {/* Header */}
          <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 md:p-8 text-white">
            <div className="flex items-center gap-3 mb-2">
              <Mail size={32} />
              <h1 className="text-2xl md:text-3xl font-bold">Outlook To-Do Agent</h1>
            </div>
            <p className="text-blue-100 text-sm md:text-base">Automatically extract tasks from your last 2 months of emails</p>
          </div>

          {/* Content */}
          <div className="p-6 md:p-8">
            {/* Setup Instructions */}
            {showSetup && (
              <div className="mb-8 p-6 bg-gradient-to-br from-blue-50 to-indigo-50 rounded-xl border-2 border-blue-200">
                <h2 className="text-xl font-bold text-gray-800 mb-4 flex items-center gap-2">
                  <AlertCircle className="text-blue-600" size={24} />
                  Quick Setup Required
                </h2>
                
                <div className="space-y-4 text-sm">
                  <div className="bg-white p-4 rounded-lg">
                    <h3 className="font-semibold text-gray-800 mb-2">Current Redirect URI:</h3>
                    <code className="bg-gray-100 px-3 py-2 rounded block text-xs break-all">
                      {window.location.origin}
                    </code>
                    <p className="text-gray-600 mt-2 text-xs">Use this URL in your Azure AD app configuration</p>
                  </div>

                  <div className="space-y-3">
                    <h3 className="font-semibold text-gray-800">Steps:</h3>
                    <ol className="list-decimal list-inside space-y-2 text-gray-700">
                      <li className="flex gap-2">
                        <span className="flex-1">Go to <a href="https://portal.azure.com" target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline inline-flex items-center gap-1">Azure Portal <ExternalLink size={12} /></a></span>
                      </li>
                      <li>Navigate to "App registrations" → "New registration"</li>
                      <li>Name: "Outlook Todo Agent"</li>
                      <li>Add the Redirect URI shown above (Web platform)</li>
                      <li>Go to "API permissions" → Add "Microsoft Graph" → "Delegated"</li>
                      <li>Add permissions: <code className="bg-gray-100 px-1">Mail.Read</code> and <code className="bg-gray-100 px-1">User.Read</code></li>
                      <li>Copy your Client ID and update the code</li>
                    </ol>
                  </div>

                  <button
                    onClick={() => setShowSetup(false)}
                    className="mt-4 text-blue-600 hover:text-blue-700 font-medium text-sm"
                  >
                    Hide Setup Instructions
                  </button>
                </div>
              </div>
            )}

            {/* Authentication & Action Buttons */}
            <div className="mb-8 flex flex-wrap gap-4">
              {!accessToken ? (
                <button
                  onClick={authenticateWithMicrosoft}
                  className="flex items-center gap-2 bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 transition-colors font-medium"
                >
                  <Mail size={20} />
                  Connect to Outlook
                </button>
              ) : (
                <button
                  onClick={fetchEmails}
                  disabled={isProcessing}
                  className="flex items-center gap-2 bg-indigo-600 text-white px-6 py-3 rounded-lg hover:bg-indigo-700 transition-colors font-medium disabled:bg-gray-400"
                >
                  {isProcessing ? <Loader size={20} className="animate-spin" /> : <Play size={20} />}
                  {isProcessing ? 'Processing...' : 'Analyze Emails'}
                </button>
              )}
              
              {!showSetup && (
                <button
                  onClick={() => setShowSetup(true)}
                  className="text-blue-600 hover:text-blue-700 font-medium"
                >
                  Show Setup Instructions
                </button>
              )}
            </div>

            {/* Error Message */}
            {error && (
              <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg flex items-start gap-3">
                <AlertCircle className="text-red-600 flex-shrink-0 mt-0.5" size={20} />
                <div className="flex-1">
                  <p className="text-red-800 font-medium">Error</p>
                  <p className="text-red-600 text-sm">{error}</p>
                </div>
              </div>
            )}

            {/* Processing Status */}
            {isProcessing && (
              <div className="mb-6 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                <p className="text-blue-800">Processing {emailCount} emails and extracting tasks...</p>
              </div>
            )}

            {/* To-Do List */}
            {todos.length > 0 && (
              <div>
                <div className="flex items-center gap-2 mb-4">
                  <CheckSquare className="text-indigo-600" size={24} />
                  <h2 className="text-xl md:text-2xl font-bold text-gray-800">
                    Extracted Tasks ({todos.length})
                  </h2>
                </div>

                <div className="space-y-4">
                  {todos.map((todo) => (
                    <div key={todo.id} className="border border-gray-200 rounded-lg p-4 hover:shadow-md transition-shadow">
                      <div className="flex items-start justify-between gap-4">
                        <div className="flex-1 min-w-0">
                          <h3 className="font-semibold text-gray-900 mb-1 break-words">{todo.title}</h3>
                          {todo.description && (
                            <p className="text-gray-600 text-sm mb-2 break-words">{todo.description}</p>
                          )}
                          <div className="flex flex-wrap gap-2 text-xs">
                            <span className={`px-2 py-1 rounded-full font-medium ${getPriorityColor(todo.priority)}`}>
                              {todo.priority || 'medium'}
                            </span>
                            {todo.dueDate && (
                              <span className="px-2 py-1 bg-purple-50 text-purple-700 rounded-full flex items-center gap-1">
                                <Calendar size={12} />
                                {new Date(todo.dueDate).toLocaleDateString()}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="mt-3 pt-3 border-t border-gray-100 text-xs text-gray-500">
                        <p className="break-words">From: <span className="font-medium">{todo.from}</span></p>
                        <p className="break-words">Email: {todo.emailSubject} ({todo.emailDate})</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default OutlookTodoAgent;
