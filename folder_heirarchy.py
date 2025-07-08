import os

# Folders to skip entirely
IGNORED_DIRS = {'.git', '__pycache__', 'env', 'venv', '.idea', '.vscode', 
               'node_modules', '.next', 'dist', 'build'}

def load_gitignore(root_path):
    """Load .gitignore entries if file exists"""
    gitignore_path = os.path.join(root_path, '.gitignore')
    ignored_files = set()
    if os.path.exists(gitignore_path):
        try:
            with open(gitignore_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        # Store just the filename/pattern, not full paths
                        ignored_files.add(os.path.basename(line.rstrip('/')))
        except (IOError, UnicodeDecodeError):
            pass  # Skip if can't read gitignore
    return ignored_files

def generate_directory_tree(start_path='.', indent='', ignored_files=None):
    """Generate directory tree with proper error handling"""
    if ignored_files is None:
        ignored_files = set()
    
    tree_output = ''
    
    try:
        items = sorted(os.listdir(start_path))
    except (PermissionError, OSError) as e:
        # Return indicator for inaccessible directories
        return indent + '└── [Permission Denied]\n'
    
    # Filter out ignored directories
    items = [item for item in items if item not in IGNORED_DIRS]
    
    for index, item in enumerate(items):
        path = os.path.join(start_path, item)
        
        # Skip if explicitly ignored in .gitignore  
        if item in ignored_files:
            continue
        
        is_last = index == len(items) - 1
        prefix = '└── ' if is_last else '├── '
        tree_output += indent + prefix + item + '\n'
        
        if os.path.isdir(path):
            extension = '    ' if is_last else '│   '
            tree_output += generate_directory_tree(path, indent + extension, ignored_files)
    
    return tree_output

def main():
    """Main function to generate and save directory tree"""
    # Your actual project directory on D: drive:
    root_dir = r"D:\Work Files\Workspace\Project 1\Flask-PQ-Extractor\Flask_V1\pq_report_app_V1"
    
    # Or use current directory if running from project folder:
    # root_dir = os.getcwd()
    
    if not os.path.exists(root_dir):
        print(f"❌ Directory not found: {root_dir}")
        return
    
    folder_name = os.path.basename(root_dir)
    print(f"Generating directory tree for: {folder_name}")
    
    # Load .gitignore entries
    ignored_files = load_gitignore(root_dir)
    
    # Generate tree with folder name as root
    tree = f"{folder_name}/\n"
    tree += generate_directory_tree(root_dir, ignored_files=ignored_files)
    
    # Save to file
    output_file = "directory_tree.txt"
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(tree)
        print(f"✅ Saved to {output_file}")
    except IOError as e:
        print(f"❌ Error saving file: {e}")
        return
    
    # Print tree (optional - comment out for large directories)
    print("\n" + tree)

if __name__ == '__main__':
    main()