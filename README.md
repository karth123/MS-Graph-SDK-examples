
# MS-Graph-SDK-examples

A repository to store examples of usage of MS-Graph SDK in python (Users and Groups). Also contains a directed tutorial on working with the SDK.

The tutorial also explores directory extensions in Users and Groups




## Initializing GraphClient

In the class Users, the GraphClient is initialized with the following code

```python
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

```

This initialization follows client credentials flow.

All clientIDs and ClientSecrets are stored in config.cfg

The app_client makes service calls to MSGraph.


## Creating necessary app registrations

App registrations are to be created for the client application (The application making calls to MSGraph), the directory extension application for users, and the directory extension application for groups.

```python
[azure]
clientId = 
clientSecret = 
tenantId = 
user_dir_app = 
user_dir_obj = 
group_dir_app = 
group_dir_obj = 
```
The clientID parameter is the clientID of the calling application. The clientSecret is its secret.

The tenantId is the ID of the tenant organization in which these apps are registered. These examples are relevant for single tenants only.

user_dir_app is the application(Client) ID of the directory extension application registered for users.

group_dir_app is the application(Client) ID of the directory extension application registered for groups.

user_dir_obj is the object ID of the directory extension application registered for users

group_dir_obj is the object iD of the directory extension application registered for groups.

Fill config.cfg with the necessary secret values to run the application.
## Work with Directory Extensions in MSGraph

Directory extensions are strongly-typed extensions on familiar directoryObject resource types. This repository contains usage examples for the directoryObjects resource types, Users and Groups. However, they can be used for other objects as well.

Only 100 properties across all types (Directory, open, schema and untyped) can be written to a directoryObject at any time.

Directory extension properties are registered from owner applications. Directory extension properties can be registered to only one directoryObject resource type at a time. Each directory extension property can receive one value of string type only.

There are other constraints with respect to directory extensions, you are encouraged to check MSLearn documentation for them.

In tenant.py

```python
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
```

The ExtensionProperty() class is the gateway to creating extension properties for a resource type. It is good practice to have different owner applications for different resource types.

Assume a property_name "Domicile" is given. When the directory extension property is created under this property name, its calling name will become extension_{application_id of owner application without hyphens}_Domicile

To improve readability of fetched directory extension properties across the repository, a function convert_key is used to remove unnecessary information from the calling name.

``` python
def convert_key(key):
            sliced_key = key.split('_', 2)[-1]
            converted_key = re.sub(r'_', ' ', sliced_key)
            return converted_key
```

this will return back Domicile when the system generated calling name is provided.

In addition to the calling name, the system also generates a unique ID for each directory extension property. This ID will be required in order to delete any properties.

Get directory extension properties for a resource.
