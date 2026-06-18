import os, sys, datetime

# LOG STARTUP - confirms which script is running
log_path = os.path.join(os.path.dirname(__file__), "_lm_startup.log")
with open(log_path, "w") as lf:
    lf.write("LOCATION MANAGER v2 BYPASS\n")
    lf.write("Script: %s\n" % __file__)
    lf.write("Time: %s\n" % datetime.datetime.now().isoformat())
    lf.write("Size: %d bytes\n" % os.path.getsize(__file__))
    lf.write("Version: 2.0 XamlReader.Load\n")

print("LM v2.0 loaded!")

SCRIPT_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib", "GUI", "Tools", "LocationManager.xaml")

with open(log_path, "a") as lf:
    lf.write("XAML_PATH: %s\n" % XAML_FILE)
    lf.write("XAML_EXISTS: %s\n" % os.path.isfile(XAML_FILE))
    if os.path.isfile(XAML_FILE):
        lf.write("XAML_SIZE: %d\n" % os.path.getsize(XAML_FILE))

if not os.path.isfile(XAML_FILE):
    import clr
    clr.AddReference("System.Windows.Forms")
    from System.Windows.Forms import MessageBox
    MessageBox.Show("XAML not found:\n" + XAML_FILE, "LM v2 Error")
else:
    import clr
    clr.AddReference("PresentationFramework")
    clr.AddReference("System")
    from System.Windows.Markup import XamlReader
    from System.IO import MemoryStream
    from System.Text import Encoding
    with open(XAML_FILE, "r") as f:
        xaml = f.read()
    stream = MemoryStream(Encoding.UTF8.GetBytes(xaml))
    try:
        window = XamlReader.Load(stream)
        with open(log_path, "a") as lf:
            lf.write("XAML_LOADED: SUCCESS\n")
        window.ShowDialog()
    except Exception as ex:
        with open(log_path, "a") as lf:
            lf.write("XAML_ERROR: %s\n" % str(ex))
            import traceback
            traceback.print_exc(file=lf)
        import clr
        clr.AddReference("System.Windows.Forms")
        from System.Windows.Forms import MessageBox
        MessageBox.Show("XAML load failed:\n" + str(ex), "LM v2 Error")
