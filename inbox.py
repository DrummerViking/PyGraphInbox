import configparser
import os
import sys
from azure.identity import InteractiveBrowserCredential, ClientSecretCredential, DeviceCodeCredential
from configparser import SectionProxy
from msgraph.core import GraphClient

class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphClient
    client_credential: ClientSecretCredential
    app_client: GraphClient
    username: str

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings["clientId"]
        tenant_id = self.settings["tenantId"]
        if "clientsecret" in list(self.settings.keys()):
            print("Logging on with Application flow.")
            secret = self.settings["clientSecret"]
            graph_scopes = self.settings["graphappscopes"].split(" ")
            self.client_credential = ClientSecretCredential(
                tenant_id, client_id, secret
            )
            self.app_client = GraphClient(
                credential=self.client_credential, scopes=graph_scopes
            )
        else:
            print("Logging on with Delegation flow.")
            graph_scopes = self.settings["graphuserscopes"].split(" ")
            self.device_code_credential = DeviceCodeCredential(
                client_id=client_id, tenant_id=tenant_id
            )
            self.user_client = GraphClient(
                credential=self.device_code_credential, scopes=graph_scopes
            )

    def get_token(self):
        if "clientsecret" in list(self.settings.keys()):
            graph_scopes = self.settings["graphappscopes"]
            access_token = self.client_credential.get_token(graph_scopes)
        else:
            graph_scopes = self.settings["graphuserscopes"]
            access_token = self.device_code_credential.get_token(graph_scopes)
        return access_token.token

    def get_inbox(self, primarysmtpaddress):
        endpoint = f"/users/{primarysmtpaddress}/mailfolders/inbox/messages"
        # Only request specific properties
        select = "from,isRead,receivedDateTime,subject"
        # Get at most 25 results
        top = 25
        # Sort by received time, newest first
        order_by = "receivedDateTime DESC"
        request_url = f"{endpoint}?$select={select}&$top={top}&$orderBy={order_by}"
        if "clientsecret" in list(self.settings.keys()):
            inbox_response = self.app_client.get(request_url)
        else:
            inbox_response = self.user_client.get(request_url)
        return inbox_response.json()

if __name__ == "__main__":
    # Load settings
    config = configparser.ConfigParser()
    this_file_path = os.path.abspath(__name__)
    BASE_DIR = os.path.dirname(this_file_path)
    config_file = os.path.join(BASE_DIR, "config-dev.cfg")
    if os.path.exists(config_file):
        print(f"Using {config_file}")
        config.read(config_file)
    else:
        config_file = os.path.join(BASE_DIR, "config.cfg")
        if not os.path.exists(config_file):
            print(f"{config_file} does not exists.")
            raise FileNotFoundError
        print(f"Using {config_file}")
        config.read(config_file)
    azure_settings = config["Azure"]
    # create graph object
    graph: Graph = Graph(azure_settings)
    # get email messages
    # checking if users file list exists in config file
    if len(sys.argv) > 1:
        users_list = sys.argv[1].split(",")
    elif "userslistfilename" in list(azure_settings.keys()):
        users_file = os.path.join(BASE_DIR, azure_settings["userslistfilename"])
        if os.path.exists(users_file):
            with open(users_file, "r") as f:
                content = f.read()
                users_list = content.split("\n")
        else:
            print("Users list file: '{0}' does not exists.".format(azure_settings["userslistfilename"]))
            raise FileNotFoundError
    else:
        print("No users file, nor email addresses as arguments passed.")
        exit()
    for user in users_list:
        print(f"Getting messages for mailbox {user}: ")
        message_page = graph.get_inbox(user)
        if "error" in message_page.keys():
            print(
                "Error occurred to get messages for user {0}. {1}".format(
                    user, message_page["error"]["message"]
                )
            )
            break
        for message in message_page["value"]:
            print("Message:", message["subject"])
            print("  From:", message["from"]["emailAddress"]["name"])
            print("  Status:", "Read" if message["isRead"] else "Unread")
            print("  Received:", message["receivedDateTime"])
        # If @odata.nextLink is present
        more_available = "@odata.nextLink" in message_page
        print("\nMore messages available?", more_available, "\n")
