from configparser import SectionProxy
from azure.identity.aio import ClientSecretCredential
from typing import List,Dict
import re
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider
)
from msgraph import GraphRequestAdapter, GraphServiceClient
from msgraph.generated.applications.get_available_extension_properties import \
    get_available_extension_properties_post_request_body
from msgraph.generated.models.extension_property import ExtensionProperty


class Tenant:
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
        auth_provider = AzureIdentityAuthenticationProvider(self.client_credential)
        self.request_adapter = GraphRequestAdapter(auth_provider)
        self.app_client = GraphServiceClient(self.request_adapter)

        # Create directory extensions for groups

    async def create_directory_extension_properties_for_groups(self,properties:list):
        object_id = self.settings['group_dir_obj']
        for property_name in properties:
            request_body = ExtensionProperty()
            request_body.name = re.sub(r'\s', '_', property_name)
            request_body.data_type = 'String'
            request_body.target_objects = (['Group', ])
            result = await self.app_client.applications.by_application_id(object_id).extension_properties.post(
                request_body)
        pass

        # Create directory extensions for users

    async def user_properties_builder_flow(self,properties:list):
        object_id = self.settings['user_dir_obj']
        for property_name in properties:
            request_body = ExtensionProperty()
            request_body.name = re.sub(r'\s', '_', property_name)
            request_body.data_type = 'String'
            request_body.target_objects = (['User', ])
            result = await self.app_client.applications.by_application_id(object_id).extension_properties.post(
                request_body)
        pass

    # Get all extension properties for a user

    async def fetch_extensions_user(self):
        application_id = self.settings['user_dir_app']
        def convert_key(key):
            sliced_key = key.split('_', 2)[-1]
            converted_key = re.sub(r'_', ' ', sliced_key)
            return converted_key
        extension_request_body = get_available_extension_properties_post_request_body.GetAvailableExtensionPropertiesPostRequestBody()
        extension_request_body.is_synced_from_on_premises = False
        result = await self.app_client.directory_objects.get_available_extension_properties.post(
            extension_request_body)
        extension_properties =[]
        for value in result.value:
            if value.name[10:42] == re.sub("-", "", application_id):
                extension_properties.append({"Name": convert_key(value.name), "ID": value.id})
        return extension_properties
    
    # Get all extension properties for a group

    async def fetch_extensions_group(self):
        application_id = self.settings['group_dir_app']

        # Helper function to make additional extension data readable.
        def convert_key(key):
            sliced_key = key.split('_', 2)[-1]
            converted_key = re.sub(r'_', ' ', sliced_key)
            return converted_key

        extension_request_body = get_available_extension_properties_post_request_body.GetAvailableExtensionPropertiesPostRequestBody()
        extension_request_body.is_synced_from_on_premises = False
        result = await self.app_client.directory_objects.get_available_extension_properties.post(
            extension_request_body)
        extension_properties =[]
        for value in result.value:
            if value.name[10:42] == re.sub("-", "", application_id):
                extension_properties.append({"Name": convert_key(value.name), "ID": value.id})
        return extension_properties
    
    # Delete extension properties for a user  (Input a list containing the ids of the extension properties to delete)

    async def delete_user_properties(self,property_ids:list):
        obj_id = self.settings['user_dir_obj']
        for property_id in property_ids:
            await self.app_client.applications.by_application_id(obj_id).extension_properties.by_extension_property_id(property_id).delete()
    
