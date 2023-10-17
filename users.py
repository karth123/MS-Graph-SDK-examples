from configparser import SectionProxy
from azure.identity.aio import ClientSecretCredential
from typing import List
from typing import Dict
import string
import random
import re
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider
)
from msgraph import GraphRequestAdapter, GraphServiceClient
from msgraph.generated.applications.get_available_extension_properties import \
    get_available_extension_properties_post_request_body
from msgraph.generated.models.password_profile import PasswordProfile
from msgraph.generated.models.user import User
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

# Authenticate and initialize Microsoft Graph Client using client credentials flow (Version 1.0.0.a12)
class Users:
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
    
    
    
    # Get all users in the tenant (Select only DisplayName, id and jobTitle)

    async def get_all_users(self):
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            select=['displayName', 'id', 'jobTitle'],
            orderby=['displayName']
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        response = await self.app_client.users.get(request_configuration=request_config)
        users = []
        for user in response.value:
            user_data = {}
            user_data['Name'] = user.display_name
            user_data['id'] = user.id
            job_title = user.job_title
            user_data["jobTitle"] = job_title
            users.append(user_data)
        return user_data
    
    # Update directory extension properties of the user
    # user_dir_app is the directory extension application ID for users
    # Ensure that the property has been added to the tenant before updating

    async def update_user(self, user_id, property_name: str, property_value: str):
        user_dir_app = self.settings['user_dir_app']
        dir_property = f"extension_{re.sub('-', '', user_dir_app)}_" + property_name.replace(" ", "_")
        request_body = User()
        additional_data = {f"{dir_property}": None}
        additional_data[f"{dir_property}"] = property_value
        request_body.additional_data = additional_data
        result = await self.app_client.users.by_user_id(user_id).patch(request_body)
        pass

    # Create a single user (Input user properties in a dictionary format. The minimum requirement is Name)

    async def user_creation_singular(self, user_properties):
        # app_id = self.settings['user_dir_app'] Optional if using directory extensions
        request_body = User()
        request_body.account_enabled = True
        display_name = user_properties['Name']
        mail = ''.join([word.capitalize() for word in display_name.split()]) + "@v2tzs.onmicrosoft.com"
        # Helper functions

        def password_generate_msft():
            characters = string.digits + string.punctuation + string.ascii_uppercase + string.ascii_lowercase
            password = ''.join(random.choice(characters) for i in range(20))
            return password

        password = password_generate_msft()
        request_body.display_name = display_name
        request_body.mail_nickname = ''.join([word.capitalize() for word in display_name.split()])
        request_body.user_principal_name = mail
        request_body.mail = mail
        password_profile = PasswordProfile()
        password_profile.force_change_password_next_sign_in = True
        password_profile.password = password
        request_body.password_profile = password_profile

        # Optional parameters (Directory extensions)
        # Ensure that the properties are defined in the tenant before removing code comments

        # request_body.job_title = user_properties["JobTitle"]  (Optional)
        # def transform_key(key):
        #     return f"extension_{re.sub('-', '', app_id)}_{re.sub(' ', '_', key)}"

        # replace_keys = lambda obj: {transform_key(k): replace_keys(v) if isinstance(v, dict) else v for k, v in
        #                             obj.items()}
        # additional_data = replace_keys(user_properties)
        # request_body.additional_data = additional_data

        result = await self.app_client.users.post(request_body)
        user_id = result.id
        return password,mail,user_id
    
    # Get all user information from the user_id (Includes directory extension data)

    async def get_user_by_id(self, id_num:str): 
        application_id = self.settings['user_dir_app']
        application_id = application_id.strip()
        extension_request_body = get_available_extension_properties_post_request_body.GetAvailableExtensionPropertiesPostRequestBody()
        extension_request_body.is_synced_from_on_premises = False
        result = await self.app_client.directory_objects.get_available_extension_properties.post(
            extension_request_body)
        extension_property_names_with_app = []
        extension_property_names = []

        # Helper function to convert the unreadable additional data into readable form
        def convert_key(key):
            sliced_key = key.split('_', 2)[-1]
            converted_key = re.sub(r'_', ' ', sliced_key)
            return converted_key

        for value in result.value:
            if value.name[10:42] == re.sub("-", "", application_id):
                extension_property_names_with_app.append(value.name)
                extension_property_names.append(convert_key(value.name))
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            select=['displayName', 'id', 'jobTitle'] + [str(value) for value in extension_property_names_with_app],
            # Sort by display name
            orderby=['displayName'],
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.app_client.users.by_user_id(id_num).get(request_config)
        user_data = {}
        user_data["Name"] = user.display_name
        user_data["jobTitle"] = user.job_title
        user_data['id'] = user.id
        # Similarly, add other properties as required
        del user.additional_data["@odata.context"]  # Removes unnecessary information in the additional data
        user_data["properties"] = user.additional_data
        return user_data
    
    # delete user from tenant

    async def delete_user(self, user_id):
        await self.app_client.users.by_user_id(user_id).delete()
        pass

    # Get all groups a user belongs to  (List of Group IDs)

    async def get_groups_of_user(self,user_id):
        group_ids = []
        result = await self.app_client.users.by_user_id(user_id).member_of.graph_group.get()
        for val in result.value:
            group_ids.append(val.id)
        return group_ids
    
    # Helper functions

    def password_generate_msft():
        characters = string.digits + string.punctuation + string.ascii_uppercase + string.ascii_lowercase
        password = ''.join(random.choice(characters) for i in range(20))
        return password

    def convert_key(key):
        sliced_key = key.split('_', 2)[-1]
        converted_key = re.sub(r'_', ' ', sliced_key)
        return converted_key
