"""
Get all Service Profiles from all defined UCS domains and export them to xls file
"""

import pandas as pd

from ucsmsdk.ucshandle import UcsHandle
from domains import ucs_domains_list


def get_sp_from_domain(domain):
    # Create a connection handle
    handle = UcsHandle(domain['ip'], domain['username'], domain['password'])
    # Login to the server
    handle.login()
    # Get all service profiles
    query = handle.query_classid(class_id="LsServer")
    # Logout from the server
    handle.logout()
    # Convert them in a list with desired attributes
    sp_list = [[sp.name, sp.type, sp.assign_state, sp.assoc_state] for sp in query]
    return sp_list


def convert_sp_list_to_excel(sp_list, xls_file_name, sheet_name):
    # convert to Pandas DF and export to excel
    df = pd.DataFrame.from_records(sp_list, columns=['SP Name', 'Type', 'Assigned State', 'Associated State'],)
    # print(df)
    df.to_excel(xls_file_name, sheet_name=sheet_name)


def main():
    for domain in ucs_domains_list:
        sp_list = get_sp_from_domain(domain)
        convert_sp_list_to_excel(sp_list=sp_list, xls_file_name='file.xls', sheet_name=domain['name'])


if __name__=='__main__':
    main()
