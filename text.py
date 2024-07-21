import config.server_cfg as config
import tableauserverclient as TSC
from datetime import datetime, timezone
import pandas as pd

# Global variables
INACTIVE_THRESHOLD = 500  # days


def tableau_signin(site_id):
    '''Sign in to Tableau server'''
    tableau_auth = TSC.PersonalAccessTokenAuth(
        token_name=config.HYPER_TOKEN_NAME,
        personal_access_token=config.HYPER_TOKEN,
        site_id=site_id
    )
    server = TSC.Server(config.HYPER_SITE_URL, use_server_version=False)
    # server.add_http_options({'verify': False})
    server.auth.sign_in(tableau_auth)
    return server

def get_users(server):
    '''Get all the users from the Tableau server'''
    all_users, pagination_item = server.users.get()
    print("INFO: Total number of users: ", pagination_item.total_available)
    return all_users

def get_sites(server):
    '''Get all the sites from the Tableau server'''
    sites_list =[]
    all_sites, pagination_item = server.sites.get()
    for site in all_sites:
        sites_list.append(site.content_url)
    print("INFO: Total number of sites: ", pagination_item.total_available)
    server.auth.sign_out()
    return sites_list

def get_inactive_users(viewer_list):
    '''Identify viewers whose last sign in > 500 days : Inactive users'''
    inactive_users = []
    for user in viewer_list:
        if user.last_login is not None:
            now = datetime.now(timezone.utc)
            days_inactive = (now - user.last_login).days
            if days_inactive > INACTIVE_THRESHOLD:
                inactive_users.append(user)
                print("Inactive user: ", user.name, " Last login: ", user.last_login, " Days inactive: ", days_inactive)
    print("INFO: Total number of inactive users: ", len(inactive_users))
    return inactive_users


def save_info(inactive_users):
    '''Save the inactive user info to a csv file'''
    parsed_users_info = []
    for user in inactive_users:
        user_info = {
            "ID": user.id,
            "Name": user.name,
            "Role": user.site_role,
            "Last Login": user.last_login,
            "Email": user.email,
            "Full Name": user.fullname
        }
        parsed_users_info.append(user_info)
    df = pd.DataFrame(parsed_users_info)
    csv_filename = "users_info.csv"
    df.to_csv(csv_filename, index=False)
    print(f"User information saved to {csv_filename}")


def deactivate_users(inactive_users, server):
    '''Deactivate the inactive users'''
    if len(inactive_users) == 0:
        print("INFO: No inactive users to deactivate.")
        return
    for user in inactive_users:
        print(f"INFO: Deactivating user: {user.name} ...", end="")
        user.site_role = "Unlicensed"
        server.users.update(user)
        print("\tdone!")
    print("INFO: Deactivation of users completed.")


def disable_minimum_site_role(server):
    '''Disable the minimum site role for all user groups'''
    all_groups, pagination_item = server.groups.get()
    for group in all_groups:
        if group.name == "All Users":
            print(f"INFO: Disabling minimum site role for group: {group.name}")
            group.minimum_site_role = None
            server.groups.update(group)


def enable_minimum_site_role(server):
    '''Enable the minimum site role for all user groups'''
    all_groups, pagination_item = server.groups.get()
    for group in all_groups:
        if group.name == "All Users":
            print(f"INFO: Enabling minimum site role for group: {group.name}")
            group.minimum_site_role = 'Viewer'  # Adjust this to the desired role
            server.groups.update(group)


if __name__ == "__main__":
    print("--**--")
    try:

        site_list = []
        site_list = get_sites()
        for site_list in site_list:
            print(f"Processing site: {site_list}")
            # Sign in to Tableau server
            server = tableau_signin(site_list)

            try:
                # Task 0: Disable minimum site role before processing users
                disable_minimum_site_role(server)

                # Task 1: Get all the users from the Tableau server
                user_list = get_users(server)
                print()

                # TASK 2: Identify users with Site role = viewer
                viewer_list = []
                for user in user_list:
                    if user.site_role == "Viewer":
                        viewer_list.append(user)
                print("INFO: Total number of users with Site role as Viewer: ", len(viewer_list))
                print()

                # TASK 3: Identify viewers  whose last sign in > 700 days : Inactive users
                inactive_users = get_inactive_users(viewer_list)
                print()

                # inactive_users = inactive_users[:1]

                # TASK 4: save the inactive user info to a csv file
                print("INFO: Saving the inactive user info to a csv file...")
                save_info(inactive_users)
                print()

                # TASK 5: Deactivate the inactive users
                deactivate_users(inactive_users)
                print()

                # TASK 6:Re-enable minimum site role after processing users
                enable_minimum_site_role(server)

            except Exception as e:
                print(f"ERROR: Failed to process site '{site_list}' - {e}")
            finally:
                server.auth.sign_out()
                print(f"INFO: Successfully signed out of Tableau server for site '{site_list}'.")

    except Exception as e:
        print("ERROR: ", e)
    finally:
        print("--**--")
