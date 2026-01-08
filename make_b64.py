# make_b64.py
import base64
path = "assets/RolesViews.xlsx"
b64 = base64.b64encode(open(path, "rb").read()).decode("ascii")
open("rolesviews_b64.txt", "w").write(b64)
print("Base64 gravado em rolesviews_b64.txt")
