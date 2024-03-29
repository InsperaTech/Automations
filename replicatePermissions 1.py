'''
This script is capable of detecting and replicating prod group permissions with dev groups in Tableau Server.
Author: Lalit Gupta

Steps:
1. Connect to Tableau Server
2. Get all the projects & groups on the server
3. For each project, get the permissions for the project
4. For each permission, check if the group is a production group
5. If the group is a production group, get the dev group
6. If the dev group does not exist, create it : TODO (Future) 
7. Apply the permissions to the dev group
8. Replicate the permissions to the workbook and datasource level
'''

import tableauserverclient as TSC

# Configurtion
server_url = 'https://prod-useast-b.online.tableau.com/'
site_name = 'b360bi'
token_name = 'myToken'
access_token = 'OQlxX9hPR+qhWzx3bOlv/Q==:SYpTs58I4M2xFCv4mcF6xbhbOzxDv1c7'
api_version = '3.22'

# ppt_location = r'C:\Workspace\Automations\content'
ppt_location = ''
img_export = "img/"
img_indexing = []


def connect_tableau():
    # Step 2: Connect to the Tableau Server
    try:
        tableau_auth = TSC.PersonalAccessTokenAuth(token_name=token_name,
                                                   personal_access_token=access_token,
                                                   site_id=site_name)
        server = TSC.Server(server_url)
        server.auth.sign_in(tableau_auth)
        print("DEBUG: Connected to Tableau Server")
        return server
    except Exception as err:
        print("ERR: Error connecting to Tableau Server")
        raise err

def get_projects(server: TSC.Server) -> dict:
    '''
    Get all the projects on the server
    
    Args:
    server: Tableau Server connection object
    '''
    # Step 3: Get all the projects
    all_projects, pagination_item = server.projects.get()
    projects = {}
    for project in all_projects:
        projects[project.id] = project
    return projects


def get_all_groups(server: TSC.Server) -> dict:
    '''
    Get all the groups on the server

    Args:
    server: Tableau Server connection object
    '''
    all_groups, pagination_item = server.groups.get()
    # print("There are {} groups on site: ".format(pagination_item.total_available))    
    groups = {}
    for group in all_groups:
        groups[group.id] = group
    return groups


def check_prod_group(group):
    '''
    Check if the group is a production group
    
    Args:
    group: Group object
    '''
    is_prod = False
    # Local condition: Check if the group name starts with 'prod'
    if group.name.lower().startswith('prod'):
        is_prod = True
    return is_prod


def get_prod_permissions(project, groups):
    '''
    Get all the permissions for the project
    
    Args:
    project: Project object
    groups: All the groups on the server
    '''
    prod_permissions = []
    for permission in project.permissions:
        if not permission.grantee.tag_name == 'group':
            continue
        group = groups[permission.grantee.id]
        if check_prod_group(group):
            prod_permissions.append(permission)
    return prod_permissions


def get_dev_group(group, group_index):
    group_name = group.name.lower()
    
    # TODO: Chnage with original logic
    if group_name.startswith('prod'):
        group_name = group_name.replace('prod', 'dev')
    
    # Check if the group already exists
    for group in group_index.values():
        if group.name.lower() == group_name:
            return group
    # TODO: If the group does not exist, create it
    return None


def get_default_permission(project, group_id, permission_type='workbook'):
    '''
    Get the default permission for the group on the project
    '''
    permissions = project.default_workbook_permissions
    if permission_type == 'datasource':
        permissions = project.default_datasource_permissions
    for permission in permissions:
        if permission.grantee.id == group_id:
            return permission.capabilities

def get_permission_rule(dev_group, permission):
    '''
    Get the permission rule for the group

    Args:
    dev_group: Group object
    permission: Permission object
    '''
    new_permission_rule = TSC.PermissionsRule(
                grantee=dev_group,
                capabilities=permission.capabilities
            )
    return new_permission_rule

def apply_permissions(server, project, dev_group, permission):
    '''
    Apply the permissions to the project

    Args:
    server: Tableau Server connection object
    project: Project object
    dev_group: Group object
    permission: Permission object
    '''
    # Step 2: Create the permission for the dev group
    print("\t- Creating permission for dev group: {}".format(dev_group.name))
    new_permission_rule = get_permission_rule(dev_group, permission)
    print(f"\t- Udapting permission for group {dev_group.name} at project level...", end='')
    server.projects.update_permission(project, [new_permission_rule])
    print(" Done")
    
    workbook_permissions = get_default_permission(project, permission.grantee.id)
    new_permission_rule = TSC.PermissionsRule(dev_group, workbook_permissions)
    print(f"\t- Udapting permission for group {dev_group.name} at workbook level...", end='')
    server.projects.update_workbook_default_permissions(project, [new_permission_rule])
    print(" Done")
    
    datasource_permissions = get_default_permission(project, permission.grantee.id, 'datasource')
    new_permission_rule = TSC.PermissionsRule(dev_group, datasource_permissions)
    print(f"\t- Udapting permission for group {dev_group.name} at datasource level...", end='')
    server.projects.update_datasource_default_permissions(project, [new_permission_rule])
    print(" Done")


def replicate_dev_permissions(server, project, valid_permisssions, groups):
    '''
    Replicate the permissions from one project to another
    
    Args:
    server: Tableau Server connection object
    project: Project object
    valid_permisssions: List of valid permissions
    groups: All the groups on the server
    '''
    print("\t- Processing: ...Replicating permissions")
    for permission in valid_permisssions:
        try:
            # Step 1: Get the group for replication
            group_name = groups[permission.grantee.id].name
            print(f'\t- Checking for the dev group for group {group_name}...', end='')
            dev_group = get_dev_group(groups[permission.grantee.id], groups)
            if not dev_group:
                print(" ERR: Dev group not found for group {}".format(permission.grantee.id))
                continue
            print(f' Found dev group for: {dev_group.name}')
            apply_permissions(server, project, dev_group, permission)
            print("\t- Permission replicated for group: {}".format(dev_group.name))

        except Exception as err:
            print(f" Err: Error replicating permissions for group {group_name}")
            raise err


def init_replicate(server: TSC.Server, all_projects: dict, all_groups: dict):
    '''
    Initialize the replication of permissions from one project to another
    
    Args:
    server: Tableau Server connection object
    all_projects: All the projects on the server
    all_groups: All the groups on the server
    '''

    # Step 4: Populate the projects with permissions
    print("Debug. Populating the projects with permissions ...", end="")
    for project_id, project in all_projects.items():
        # print("\nProject: {}".format(project.name))
        server.projects.populate_permissions(project)
        server.projects.populate_workbook_default_permissions(project)
        server.projects.populate_datasource_default_permissions(project)
    print("Done")

    print("Debug: Starting the process of replicating permissions from dev to prod")
    for project_id, project in all_projects.items():
        # STEP 1: if Valid project with valid permissions
        valid_permisssions = get_prod_permissions(project, all_groups)
        if len(valid_permisssions) == 0:
            continue
        print("-------------------------------------------------------------------------------------------")
        print(f"Debug: {project.name} has {len(valid_permisssions)} valid permissions.")
        # STEP 2: For each permission, run replicate_dev_permissions
        replicate_dev_permissions(server, project, valid_permisssions, all_groups)   
    


if __name__ == "__main__":

    # Step 1: Connect to Tableau Server
    server = connect_tableau()
    server.version = '3.22'

    # Step 2: Get all the projects and Groups
    all_projects, all_groups = get_projects(server), get_all_groups(server)

    # Step 3: Initialize the replication
    init_replicate(server, all_projects, all_groups)

