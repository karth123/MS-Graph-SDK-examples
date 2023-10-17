import configparser
from configparser import SectionProxy
from azure.identity.aio import ClientSecretCredential
from typing import List,Dict
import re

from azure.core.exceptions import AzureError
#from azure.cosmos import CosmosClient, PartitionKey

from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider
)
from msgraph import GraphRequestAdapter, GraphServiceClient
from msgraph.generated.applications.get_available_extension_properties import \
    get_available_extension_properties_post_request_body
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
from msgraph.generated.models.group import Group
from msgraph.generated.models.reference_create import ReferenceCreate
from msgraph.generated.models.user import User
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

config = configparser.ConfigParser()
config.read(['config.cfg', 'config.dev.cfg'])
azure_settings = config['azure']

class Groups:
    settings: SectionProxy
    client_credential: ClientSecretCredential
    request_adapter: GraphRequestAdapter
    app_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        client_secret = self.settings['clientSecret']
        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        auth_provider = AzureIdentityAuthenticationProvider(self.client_credential)  # type: ignore
        self.request_adapter = GraphRequestAdapter(auth_provider)
        self.app_client = GraphServiceClient(self.request_adapter)

    # Get all groups in the tenant  (Only DisplayName and id)

    async def get_all_groups(self):
        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            select=['displayName', 'id'],    # You can select other properties as well
            orderby=['displayName']   # Order by display name
        )
        request_config = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        groups = await self.app_client.groups.get(request_configuration=request_config)
        group_details = []
        for group in groups.value:
            group_data = {}
            group_data['displayName'] = group.display_name
            group_data['id'] = group.id
            # Similarly, add other properties
            group_details.append(group_data)
        return groups
    
    # Get all information about a group from its id (Including all members in the group)

    async def get_group_by_id(self,group_id):
        application_id = self.settings['group_dir_app']
        application_id = application_id.strip()
        extension_request_body = get_available_extension_properties_post_request_body.GetAvailableExtensionPropertiesPostRequestBody()
        extension_request_body.is_synced_from_on_premises = False
        result = await self.app_client.directory_objects.get_available_extension_properties.post(
            extension_request_body)
        extension_property_names_with_app = []
        extension_property_names = []
        # Helper function to make the additional data readable.
        def convert_key(key):
            sliced_key = key.split('_', 2)[-1]
            converted_key = re.sub(r'_', ' ', sliced_key)
            return converted_key

        for value in result.value:
            if value.name[10:42] == re.sub("-", "", application_id):
                extension_property_names_with_app.append(value.name)
                extension_property_names.append(convert_key(value.name))
        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            # Selects displayName,id and description, along with all directory extension property values (You can add other attributes in select)
            select=['displayName', 'id', 'description'] + [str(value) for value in extension_property_names_with_app],
            orderby=['displayName'],
        )
        request_config = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        group = await self.app_client.groups.by_group_id(group_id).get(request_config)
        member_query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            select = ['displayName','id'],

        )
        # Gets all members of the group
        member_request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(query_parameters=member_query_params,headers = {
		'ConsistencyLevel' : "eventual",
})

        members = await self.app_client.groups.by_group_id(group_id).members.get(request_configuration=member_request_config)
        # collects only relevant information and displays it.
        group_data = {}
        group_data["Name"] = group.display_name
        group_data['id'] = group.id
        group_data['description'] = group.description
        # Similarly, add other attributes
        group_data["properties"] = group.additional_data
        group_data["members"] = members.value
        return group_data
    
    # Create an M365 unified group from a dictionary 'group_details'. group_details must contain atleast "displayName", "Description"

    async def create_group(self, group_details):
        app_id = self.settings['group_dir_app']
        group_name = group_details['displayName']
        #group_properties = group_details["Properties"]   Optional 
        request_body = Group()
        request_body.display_name = group_name
        request_body.mail_enabled = True
        request_body.mail_nickname = re.sub(' ','_',group_details['Name'])
        request_body.security_enabled = False
        request_body.group_types = ["Unified",]
        request_body.description = group_details['Description']
        def transform_key(key):
            return f"extension_{re.sub('-', '', app_id)}_{re.sub(' ', '_', key)}"

        replace_keys = lambda obj: {transform_key(k): replace_keys(v) if isinstance(v, dict) else v for k, v in
                                    obj.items()}
        #additional_data = replace_keys(group_properties)  Optional
        #request_body.additional_data = additional_data  Optional
        result = await self.app_client.groups.post(request_body)
        group_details['Mail'] = result.mail
        group_details['SecurityEnabled'] = result.security_enabled
        group_details['id'] = result.id
        print("Creation of the group in Azure AD was successful")
        return group_details
    
    # Add users to M365 Group. Provide list of user_ids and one group_id

    async def add_users_to_group(self,user_ids:list,group_id:str,user:User):
        for user_id in user_ids:
            request_body = ReferenceCreate()
            request_body.odata_id = f"https://graph.microsoft.com//v1.0//directoryObjects//{user_id}"
            await self.app_client.groups.by_group_id(group_id).members.ref.post(request_body)

    # Remove user from group

    async def remove_user_from_group(self, user_id, group_id):
        await self.app_client.groups.by_group_id(group_id).members.by_directory_object_id(user_id).ref.delete()

    # Update group properties (Directory extension properties) Ensure the property exists in the tenant

    async def update_group_by_id(self, group_id, property_name: str, property_value: str):
        group_dir_app = self.settings['group_dir_app']
        dir_property = f"extension_{re.sub('-', '', group_dir_app)}_" + property_name.replace(" ", "_")
        request_body = Group()
        additional_data = {f"{dir_property}": None}
        additional_data[f"{dir_property}"] = property_value
        request_body.additional_data = additional_data
        result = await self.app_client.groups.by_group_id(group_id).patch(request_body)
        pass

    # Delete group

    async def delete_group_by_id(self,group_id:str):
        await self.app_client.groups.by_group_id(group_id).delete()

    # Get users of group (id, displayName, mail)

    async def get_users_of_group(self,group_id:str):     
        user_info = []
        result = await self.app_client.groups.by_group_id(group_id).members.get()
        for value in result.value:
            user_data = { "id": value._id, "displayName": value._display_name, "mail": value._mail}
            user_info.append(user_data)
        return user_info
    
