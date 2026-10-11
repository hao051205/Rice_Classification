import urllib.request
import json

print("=== TEST 1: PREDICT SINGLE ===")
payload = json.dumps({
    "Area": 13000, "Perimeter": 460.0, "Major_Axis_Length": 190.0,
    "Minor_Axis_Length": 87.0, "Eccentricity": 0.89, "Convex_Area": 13300, "Extent": 0.65
}).encode("utf-8")

req = urllib.request.Request("http://127.0.0.1:5000/predict", data=payload, headers={"Content-Type": "application/json"})
with urllib.request.urlopen(req) as resp:
    res = json.loads(resp.read().decode("utf-8"))
    print("Prediction:", res["prediction"])
    print("Probabilities:", res["probabilities"])
    print("Confidence:", res["confidence"])
    print("Is Uncertain:", res.get("is_uncertain"))
    if res.get("is_uncertain"):
        print("Message:", res.get("uncertainty_message"))

print("\n=== TEST 2: DOWNLOAD SAMPLE CSV ===")
req2 = urllib.request.Request("http://127.0.0.1:5000/download-sample-csv")
with urllib.request.urlopen(req2) as resp:
    content = resp.read().decode("utf-8")
    print("Lines received:", len(content.splitlines()))

print("\n=== TEST 3: PREDICT BATCH CSV ===")
boundary = "----TestBoundary12345"
with open("data/sample_batch.csv", "rb") as f:
    csv_bytes = f.read()

part1 = f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"sample_batch.csv\"\r\nContent-Type: text/csv\r\n\r\n".encode("utf-8")
part2 = f"\r\n--{boundary}--\r\n".encode("utf-8")
body = part1 + csv_bytes + part2

req3 = urllib.request.Request(
    "http://127.0.0.1:5000/predict-batch",
    data=body,
    headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
)
with urllib.request.urlopen(req3) as resp:
    res_batch = json.loads(resp.read().decode("utf-8"))
    print("Total samples:", res_batch["total_samples"])
    print("Cammeo count:", res_batch["cammeo_count"])
    print("Osmancik count:", res_batch["osmancik_count"])
    print("Uncertain count:", res_batch["uncertain_count"])
    print("First 3 results:")
    for s in res_batch["samples"][:3]:
        print(f"  #{s['index']}: {s['prediction']} (Cammeo: {s['prob_cammeo']}, Osmancik: {s['prob_osmancik']}) - {s['status']}")

print("\n=== ALL TESTS PASSED! ===")
