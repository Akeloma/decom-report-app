import toxic_general
import flt_general
import toxic_detailed
import flt_detailed

def run_all(filename):
    print("🚀 Starting all reports...")
    toxic_general.main(filename)
    flt_general.main(filename)
    toxic_detailed.main(filename)
    flt_detailed.main(filename)
    print("✅ All reports completed!")
