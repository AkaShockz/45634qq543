import tkinter as tk
import traceback
import sys
import os

def create_error_window(error_message):
    """Create a simple error window to display the error message"""
    error_root = tk.Tk()
    error_root.title("Vehicle Transport Parser - Error")
    error_root.geometry("600x400")
    
    frame = tk.Frame(error_root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    header = tk.Label(
        frame, 
        text="Error Starting Application",
        font=("Arial", 14, "bold"),
        fg="#F44336"
    )
    header.pack(pady=10)
    
    details = tk.Text(
        frame,
        wrap=tk.WORD,
        height=20,
        width=70
    )
    details.insert(tk.END, error_message)
    details.pack(fill=tk.BOTH, expand=True, pady=10)
    
    close_btn = tk.Button(
        frame,
        text="Close",
        command=error_root.destroy,
        padx=15,
        pady=5
    )
    close_btn.pack(pady=10)
    
    error_root.mainloop()

def main():
    try:
        # Check if we can import all required modules
        missing_modules = []
        
        try:
            import holidays
        except ImportError:
            missing_modules.append("holidays")
        
        # If there are missing modules, show an error
        if missing_modules:
            error = f"Missing required Python modules: {', '.join(missing_modules)}\n\n"
            error += "Please install them using pip:\n"
            for module in missing_modules:
                error += f"pip install {module}\n"
            create_error_window(error)
            return
            
        # Try to import the main module
        try:
            # Add current directory to path if not already there
            current_dir = os.path.dirname(os.path.abspath(__file__))
            if current_dir not in sys.path:
                sys.path.insert(0, current_dir)
                
            from vehicle_transport_parser import VehicleTransportApp
            
            # Create the main application
            root = tk.Tk()
            app = VehicleTransportApp(root)
            root.mainloop()
            
        except Exception as e:
            error = "Error importing or starting the main application:\n\n"
            error += f"{str(e)}\n\n"
            error += "Traceback:\n"
            error += traceback.format_exc()
            create_error_window(error)
            
    except Exception as e:
        # If even our error handler fails, print to console
        print(f"Critical error: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main() 