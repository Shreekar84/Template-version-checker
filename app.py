from detector import TemplateDetector

detector = TemplateDetector("master.xlsx")

result = detector.detect("sample_old.docx")

print("\n====================================")
print("      TEMPLATE DETECTION RESULT")
print("====================================")

if "error" in result:
    print(result["error"])
else:
    print(f"Template ID      : {result['template_id']}")
    print(f"Document Version : {result['document_version']}")
    print(f"Latest Version   : {result['latest_version']}")
    print(f"Status           : {result['status']}")
    print(f"Message          : {result['message']}")

print("====================================")