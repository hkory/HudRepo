# CEATabulationCleaner
# By Hud & GPT

def CEATabCleaner(input_file_path, output_file_path):
    # Open the original text file for reading
    with open(input_file_path, 'r') as input_file:
        # Read all lines from the file
        lines = input_file.readlines()

    # Apply modifications

    # 2. Delete the first 6 lines
    lines = lines[6:]

    # 3. Delete the last 4 lines
    lines = lines[:-4]

    # 4. Delete the first 2 characters in all lines
    lines = [line[2:] for line in lines]

    # 5. Replace every 12th character with a comma
    lines = [''.join([char if (index + 1) % 12 != 0 else ',' for index, char in enumerate(line)]) for line in lines]

    # 6. Delete the first character in all lines
    lines = [line[1:] for line in lines]

    # Open a new text file for writing
    with open(output_file_path, 'w') as output_file:
        # Write the modified content to the new text file with a newline every 96 characters
        for line in lines:
            for i in range(0, len(line), 96):
                output_file.write(line[i:i + 96] + '\n')