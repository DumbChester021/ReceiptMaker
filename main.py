import datetime, re
d = {"D": [], "T": [], "N": [], "PN": [], "PP": [], "Q": [], "TP": []}
with open("input.txt", "r") as f:
    for l in f:
        l = l.strip() 
        if l.startswith("Date:"):
            d["D"].append(l)
        elif l.startswith("Time:"):
            d["T"].append(l)
        elif l.startswith("Name:"):
            d["N"].append(l)
        elif l.startswith("Product Name:"):
            d["PN"].append(l)
        elif l.startswith("Product Price:"):
            d["PP"].append(l)
        elif l.startswith("Quantity:"):
            d["Q"].append(l)
        elif l.startswith("Total Price:"):
            d["TP"].append(l)

o = 0
p = {}
n = {}

for i in range(len(d["N"])):
    na = d["N"][i].split(':')[1].strip()
    pn = d["PN"][i].split(':')[1].strip()
    pp = float(re.search(r'\d+\.\d+', d["PP"][i]).group())
    q = int(d["Q"][i].split(':')[1])
    tp = float(re.search(r'\d+\.\d+', d["TP"][i]).group())
    o += q
    if pn in p:
        p[pn]["q"] += q
        p[pn]["tp"] += tp
    else:
        p[pn] = {"q":q, "tp":tp}
    if na in n:
        n[na]["o"].append({"pn":pn, "q":q, "tp":tp})
        n[na]["ts"] += tp
    else:
        n[na] = {"o":[ {"pn":pn, "q":q, "tp":tp} ], "ts":tp}

now = datetime.datetime.now()
fname = f"orders({now.strftime('%Y-%m-%d %H-%M-%S')}).txt"
with open(fname, "w") as f:
    f.write(f"Total Orders: {o}\n\n")
    for pn, v in p.items():
        f.write(f"{pn}: {v['q']} - {v['tp']}\n")
    f.write("\n")
    total_price = sum([v['ts'] for na, v in n.items()])
    f.write(f"\nTotal Price: {total_price}\n")
    for na, v in n.items():
        f.write("\n--------------------------\n")
        f.write(f"Name: {na}\nOrders:\n")
        for o in v["o"]:
            f.write(f" {o['pn']} - x{o['q']} - {o['tp']}\n")
        f.write(f"Total Spent: {v['ts']}\n")
