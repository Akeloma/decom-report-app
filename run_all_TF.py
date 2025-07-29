import toxic_general
import flt_general
import toxic_detailed
import flt_detailed

def run_all():
    print("🚀 Starting all reports...")
    toxic_general.main()
    flt_general.main()
    toxic_detailed.main()
    flt_detailed.main()
    print("✅ All reports completed!")
