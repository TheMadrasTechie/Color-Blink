import cv2

image_path = "sundar_face.png"
scale_factor = 10  # reduce size by 10x

img = cv2.imread(image_path)
h, w = img.shape[:2]

new_w = w // scale_factor
new_h = h // scale_factor

resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)

cv2.imwrite("template_3_reduced.jpg", resized)
print(f"Saved → template_3_reduced.jpg ({new_w} x {new_h})")
