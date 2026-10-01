import openvino as ov

core = ov.Core()
print("Available Devices:")
for device in core.available_devices:
    print(f"- {device}")
