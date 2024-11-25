import sys
import csv
from robot import run

def modify_robot_file(robot_file_path, to_date, from_date):
    # Read the original content
    with open(robot_file_path, 'r') as file:
        lines = file.readlines()
    
    # Define replacement strings
    to_date_replacement = f'${{TO_DATE}}\t{to_date}\n'
    from_date_replacement = f'${{FROM_DATE}}\t{from_date}\n'
    
    # Initialize flags to ensure only the first occurrence is replaced
    to_date_replaced = False
    from_date_replaced = False

    # Modify lines
    new_lines = []
    for line in lines:
        if '${TO_DATE}' in line and not to_date_replaced:
            new_lines.append(to_date_replacement)
            to_date_replaced = True
        elif '${FROM_DATE}' in line and not from_date_replaced:
            new_lines.append(from_date_replacement)
            from_date_replaced = True
        else:
            new_lines.append(line)
    
    # Write the modified content back to the file
    with open(robot_file_path, 'w') as file:
        file.writelines(new_lines)
    
    # Print modified content
    print("Modified content:")
    with open(robot_file_path, 'r') as file:
        print(file.read())
    
    return lines

def restore_robot_file(robot_file_path, original_lines):
    # Write the original content back to the file
    with open(robot_file_path, 'w') as file:
        file.writelines(original_lines)
    
    # Print restored message
    print("The .robot file has been restored to its original state.")

def execute_robot_for_dates(robot_file_path, dates_csv_path):
    # Read the original content
    with open(robot_file_path, 'r') as file:
        original_lines = file.readlines()

    # Read dates from the CSV file
    with open(dates_csv_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        
        for row in reader:
            from_date = row['FROM_DATE']
            to_date = row['TO_DATE']
            
            print(f"Running Robot Framework for FROM_DATE={from_date} and TO_DATE={to_date}")
            
            # Modify the .robot file with the current set of dates
            modify_robot_file(robot_file_path, to_date, from_date)
            
            try:
                # Run Robot Framework with the modified file
                run(robot_file_path)
            finally:
                # Restore the original .robot file content
                restore_robot_file(robot_file_path, original_lines)
                print(f"Completed execution for FROM_DATE={from_date} and TO_DATE={to_date}")
    
    print("All date sets have been processed.")

if __name__ == "__main__":
    robot_file_path = 'C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\QB2.robot'
    dates_csv_path = 'C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\dates.csv'
    
    # Execute the Robot Framework file for each date set in the CSV
    execute_robot_for_dates(robot_file_path, dates_csv_path)
