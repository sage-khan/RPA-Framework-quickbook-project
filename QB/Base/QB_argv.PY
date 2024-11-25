import sys

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

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: run_robot.py <TO_DATE> <FROM_DATE>")
        sys.exit(1)

    to_date = sys.argv[1]
    from_date = sys.argv[2]
    
    robot_file_path = 'C:\\Users\\virtuous\\Desktop\\Work\\Automation\\QB\\QB2.robot'
    
    # Read the original content
    with open(robot_file_path, 'r') as file:
        original_lines = file.readlines()
    
    # Modify the .robot file
    modify_robot_file(robot_file_path, to_date, from_date)
    
    try:
        # Run Robot Framework with the modified file
        from robot import run
        run(robot_file_path)
    finally:
        # Restore the original .robot file content
        restore_robot_file(robot_file_path, original_lines)
