import ctypes
from ctypes import wintypes
import os

# Constants for GetOpenFileName
OFN_ALLOWMULTISELECT = 0x00000200
OFN_EXPLORER = 0x00080000
OFN_FILEMUSTEXIST = 0x00001000
OFN_PATHMUSTEXIST = 0x00000800

class OPENFILENAME(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPCWSTR),
    ]

def ask_open_filename(title="Select File", filter_str="All Files\0*.*\0"):
    """
    Opens a single file selection dialog.
    filter_str format: "Description\0*.ext\0Description2\0*.ext2\0"
    Example: "Text Files\0*.txt\0All Files\0*.*\0"
    """
    max_path = 4096
    file_buffer = ctypes.create_unicode_buffer(max_path)
    
    ofn = OPENFILENAME()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
    ofn.hwndOwner = 0
    ofn.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    ofn.nMaxFile = max_path
    ofn.lpstrFilter = filter_str
    ofn.nFilterIndex = 1
    ofn.lpstrTitle = title
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST
    
    if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return file_buffer.value
    return ""

def ask_open_filenames(title="Select Files", filter_str="All Files\0*.*\0"):
    """
    Opens a multi-file selection dialog.
    Returns a list of absolute paths.
    """
    max_path = 65536
    file_buffer = ctypes.create_unicode_buffer(max_path)
    
    ofn = OPENFILENAME()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
    ofn.hwndOwner = 0
    ofn.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    ofn.nMaxFile = max_path
    ofn.lpstrFilter = filter_str
    ofn.nFilterIndex = 1
    ofn.lpstrTitle = title
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT
    
    if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        # Read raw bytes from the buffer using string_at to handle multiple nulls
        raw_bytes = ctypes.string_at(ctypes.byref(file_buffer), ctypes.sizeof(file_buffer))
        full_content = raw_bytes.decode('utf-16-le')
        
        # Split by null characters
        parts = full_content.split('\0')
        
        cleaned_parts = []
        for p in parts:
            if not p:
                break # Stop at first empty string (double null)
            cleaned_parts.append(p)
            
        if not cleaned_parts:
            return []
            
        first_item = cleaned_parts[0]
        
        if len(cleaned_parts) == 1:
            return [first_item]
        else:
            # First is directory, rest are filenames
            directory = first_item
            files = cleaned_parts[1:]
            return [os.path.join(directory, f) for f in files]
            
    return []

if __name__ == "__main__":
    pass
