# Define the scopes required for the Microsoft Graph API
scopes = ["https://graph.microsoft.com/.default"]

# Initialize a lock for thread-safety
lock = Lock()