import cv2
import argparse

def reduce_image_size_argparse():
    parser = argparse.ArgumentParser(description="Reduce image size by a scale factor")
    parser.add_argument("--image_path", required=True, help="Path to the input image")
    parser.add_argument("--scale_factor", type=int, default=10, help="Reduction scale factor (default: 10)")
    parser.add_argument("--output", default=None, help="Output filename (optional)")
    args = parser.parse_args()

    img = cv2.imread(args.image_path)
    if img is None:
        raise FileNotFoundError(f"Image not found: {args.image_path}")

    h, w = img.shape[:2]
    new_w = w // args.scale_factor
    new_h = h // args.scale_factor

    resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)

    output_name = args.output if args.output else f"reduced_{args.scale_factor}x_{args.image_path}"
    cv2.imwrite(output_name, resized)

    print(f"Saved → {output_name} ({new_w} x {new_h})")

def reduce_image_size(image_path: str, scale_factor: int = 10, output_name: str = None):
    """
    Reduces image dimensions by the given scale factor.
    Example: 1000x1000 → 100x100 if scale_factor = 10
    """
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"Image not found: {image_path}")

    h, w = img.shape[:2]
    new_w = w // scale_factor
    new_h = h // scale_factor

    resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)

    if output_name is None:
        output_name = f"reduced_{scale_factor}x_{image_path}"

    cv2.imwrite(output_name, resized)
    print(f"Saved → {output_name} ({new_w} x {new_h})")

    return output_name