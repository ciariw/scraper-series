import json
import Functions
from Login import initiate
from Login import login
import time

with open("login.json","r") as z:
    usr, pw = json.load(z).values()
search_depth = 5
company_list = []
feature_list = ["Job Listings", "Connection Info", "Company Employee list"]

# I should outline my objectives before starting the code so that i have a framework to work with
# 1) Log in to linkedin. If the login details are wrong, throw an exception and wait for them to log in manually.
# 2) Once logged in, based on what was selected in the feature list, gather information and save data to a CSV or json
#

def main():
    try:
        # Sorry for the unreadability
        ftr = feature_list[int(input(f"Select From the list of Features {feature_list}: "
                                 f"{[a for a in range(1,len(feature_list)+1)]}\n"))-1]
    except IndexError:
        print("You did not select a valid option!")
        main()
    except ValueError:
        print("You did not select a valid option!")
        main()
    driver = initiate("https://www.linkedin.com/")
    if len(usr) * len(pw) != 0:
        login(driver, usr, pw)
    else:
        raise Exception("Your have an issue with your login credentials")
        return

    Functions.new_session(driver, ftr)
    return

if __name__ == '__main__':
    main()
