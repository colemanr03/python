This program takes input from a user and then calls two APIs that are supplied by Magicplan
https://cloud.magicplan.app/api/v2/workgroups/plans - returns a json file with information about all of the plans created for the Customer
https://cloud.magicplan.app/api/v2/plans/forms - returns a json file with all of the Custom Form information for a specific plan
This program converts the "plans" json file information into a csv file for all of the plans selected. (User can filter by creation date)
This program also creates a csv file with all of the Custom Form Data returnd by the the "forms" API for each of the selected plans
