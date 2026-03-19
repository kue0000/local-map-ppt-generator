#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate PowerPoint presentation with China and province vector maps.
Uses SVG files inserted via COM automation (PowerPoint 2019+ supports SVG).
"""

import json
import os
import sys
import requests
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
import colorsys

if sys.version_info >= (3, 7):
    import io

    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(encoding="utf-8")

matplotlib.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Arial Unicode MS",
]
matplotlib.rcParams["axes.unicode_minus"] = False

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MAPS_DIR = os.path.join(SCRIPT_DIR, "maps")
CHINA_ADCODE = "100000"
GEO_API_BASE = "https://geo.datav.aliyun.com/areas_v3/bound"

COLOR_PALETTE = [
    "#FF6B6B",
    "#4ECDC4",
    "#45B7D1",
    "#96CEB4",
    "#FFEAA7",
    "#DDA0DD",
    "#98D8C8",
    "#F7DC6F",
    "#BB8FCE",
    "#FF9FF3",
    "#54A0FF",
    "#5F27CD",
    "#00D2D3",
    "#FF6B81",
    "#7BED9F",
    "#A29BFE",
    "#FD79A8",
    "#FDCB6E",
    "#6C5CE7",
    "#E17055",
    "#00B894",
    "#0984E3",
    "#6C5B7B",
    "#F8B500",
    "#B8E994",
    "#E77F67",
    "#786FA6",
    "#F19066",
    "#63CDDA",
    "#BDC581",
    "#78C1AD",
    "#F3A683",
    "#778BEB",
    "#CF6A87",
    "#E77F67",
]


def generate_colors(n):
    if n <= len(COLOR_PALETTE):
        return COLOR_PALETTE[:n]
    colors = []
    for i in range(n):
        hue = i / n
        sat = 0.6 + (i % 3) * 0.15
        val = 0.7 + (i % 4) * 0.075
        rgb = colorsys.hsv_to_rgb(hue, sat, val)
        colors.append(
            "#{:02x}{:02x}{:02x}".format(
                int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255)
            )
        )
    return colors


def fetch_geojson(adcode, full=True):
    suffix = "_full" if full else ""
    url = f"{GEO_API_BASE}/{adcode}{suffix}.json"
    print(f"Fetching: {url}")
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching {adcode}: {e}")
        return None


def get_coordinates(geom):
    coords = []
    if geom["type"] == "Polygon":
        coords = [geom["coordinates"]]
    elif geom["type"] == "MultiPolygon":
        coords = geom["coordinates"]
    return coords


def get_bounds(all_coords):
    all_x, all_y = [], []
    for polygon in all_coords:
        for ring in polygon:
            for point in ring:
                all_x.append(point[0])
                all_y.append(point[1])
    if not all_x or not all_y:
        return 0, 1, 0, 1
    return min(all_x), max(all_x), min(all_y), max(all_y)


def draw_geometry(ax, coords, color, edge_color="black", alpha=0.7, linewidth=0.8):
    for polygon in coords:
        for ring in polygon:
            if len(ring) < 3:
                continue
            ring = np.array(ring)
            ax.fill(
                ring[:, 0],
                ring[:, 1],
                color=color,
                edgecolor=edge_color,
                alpha=alpha,
                linewidth=linewidth,
            )


def create_map_svg(data, svg_path, title, color="#4ECDC4"):
    all_coords = []
    if "features" in data:
        for feature in data["features"]:
            coords = get_coordinates(feature["geometry"])
            all_coords.extend(coords)
    elif "geometry" in data:
        all_coords = get_coordinates(data["geometry"])

    if not all_coords:
        return None

    min_x, max_x, min_y, max_y = get_bounds(all_coords)
    width, height = max_x - min_x, max_y - min_y

    if width == 0 or height == 0:
        return None

    # Use larger figure size for better quality in PPT
    fig_width = 20  # Increased from 10 for larger output
    fig_height = fig_width * (height / width)
    if fig_height > 24:
        fig_height = 24
        fig_width = fig_height * (width / height)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height), facecolor="white")
    ax.set_aspect("equal")
    ax.axis("off")

    margin = max(width, height) * 0.05
    ax.set_xlim(min_x - margin, max_x + margin)
    ax.set_ylim(min_y - margin, max_y + margin)

    draw_geometry(ax, all_coords, color, linewidth=1.0)

    plt.tight_layout(pad=0.1)
    fig.savefig(
        svg_path,
        format="svg",
        bbox_inches="tight",
        facecolor="white",
        edgecolor="none",
        pad_inches=0.05,
    )
    plt.close(fig)

    print(f"Created: {svg_path}")
    return svg_path


def create_china_overview(provinces_data, svg_path):
    all_coords = []
    for name, data in provinces_data.items():
        if "features" in data:
            for feature in data["features"]:
                coords = get_coordinates(feature["geometry"])
                all_coords.extend(coords)
        elif "geometry" in data:
            coords = get_coordinates(data["geometry"])
            all_coords.extend(coords)

    if not all_coords:
        return None

    min_x, max_x, min_y, max_y = get_bounds(all_coords)
    width, height = max_x - min_x, max_y - min_y

    # Use larger figure size for better quality in PPT
    fig_width = 28  # Increased for larger output
    fig_height = fig_width * (height / width)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height), facecolor="white")
    ax.set_aspect("equal")
    ax.axis("off")

    margin = max(width, height) * 0.02
    ax.set_xlim(min_x - margin, max_x + margin)
    ax.set_ylim(min_y - margin, max_y + margin)

    colors = generate_colors(len(provinces_data))

    for idx, (name, data) in enumerate(sorted(provinces_data.items())):
        color = colors[idx % len(colors)]
        coords = []
        if "features" in data:
            for feature in data["features"]:
                coords.extend(get_coordinates(feature["geometry"]))
        elif "geometry" in data:
            coords = get_coordinates(data["geometry"])

        draw_geometry(ax, coords, color, linewidth=0.5)

    plt.tight_layout(pad=0.1)
    fig.savefig(
        svg_path,
        format="svg",
        bbox_inches="tight",
        facecolor="white",
        edgecolor="none",
        pad_inches=0.05,
    )
    plt.close(fig)

    print(f"Created: {svg_path}")
    return svg_path


def create_pptx_with_com(svg_files, output_path):
    """Create PowerPoint with SVG files using Windows COM."""
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 required. Run: pip install pywin32")
        return False

    print("Creating PowerPoint with SVG vectors via COM...")

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    presentation = ppt_app.Presentations.Add()

    slide_width = presentation.PageSetup.SlideWidth
    slide_height = presentation.PageSetup.SlideHeight

    print(f"  Slide size: {slide_width} x {slide_height} EMUs")
    print(f"  80% size: {int(slide_width * 0.8)} x {int(slide_height * 0.8)} EMUs")

    blank_layout = presentation.SlideMaster.CustomLayouts.Item(7)

    # Cover slide
    slide = presentation.Slides.AddSlide(1, blank_layout)

    title_shape = slide.Shapes.AddTextbox(
        1, 30, int(slide_height * 0.3), slide_width - 60, 800
    )
    tf = title_shape.TextFrame
    tf.TextRange.Text = "中华人民共和国行政区划图集"
    tf.TextRange.Font.Size = 44
    tf.TextRange.Font.Bold = True

    subtitle_shape = slide.Shapes.AddTextbox(
        1, 30, int(slide_height * 0.45), slide_width - 60, 400
    )
    tf = subtitle_shape.TextFrame
    tf.TextRange.Text = "China Administrative Division Maps"
    tf.TextRange.Font.Size = 22

    info_shape = slide.Shapes.AddTextbox(
        1, 30, int(slide_height * 0.6), slide_width - 60, 300
    )
    tf = info_shape.TextFrame
    tf.TextRange.Text = "SVG矢量格式 - 右键取消组合可编辑"
    tf.TextRange.Font.Size = 14

    slide_num = 2

    for name, svg_path in svg_files.items():
        if not os.path.exists(svg_path):
            print(f"Warning: {svg_path} not found")
            continue

        print(f"  Adding slide {slide_num}: {name}")

        slide = presentation.Slides.AddSlide(slide_num, blank_layout)
        slide_num += 1

        # Title
        title_shape = slide.Shapes.AddTextbox(1, 30, 20, slide_width - 60, 400)
        tf = title_shape.TextFrame
        tf.TextRange.Text = name
        tf.TextRange.Font.Size = 24
        tf.TextRange.Font.Bold = True

        # Insert SVG - 80% of slide size, preserve aspect ratio
        abs_path = os.path.abspath(svg_path)

        # 80% of slide dimensions
        available_width = int(slide_width * 0.8)
        available_height = int(slide_height * 0.8)

        try:
            # Insert SVG at original size first
            shape = slide.Shapes.AddPicture(abs_path, False, True, 0, 0)

            # Get original dimensions
            original_width = shape.Width
            original_height = shape.Height
            original_ratio = original_width / original_height

            # Calculate scaled dimensions to fit 80% of slide
            available_ratio = available_width / available_height

            if original_ratio > available_ratio:
                # Image is wider - fit by width
                new_width = available_width
                new_height = available_width / original_ratio
            else:
                # Image is taller - fit by height
                new_height = available_height
                new_width = available_height * original_ratio

            # Apply new size (preserving aspect ratio)
            shape.Width = int(new_width)
            shape.Height = int(new_height)

            # Center on slide
            shape.Left = int((slide_width - new_width) / 2)
            shape.Top = int((slide_height - new_height) / 2)

            print(
                f"    Inserted: {abs_path} (size: {int(new_width)}x{int(new_height)})"
            )
        except Exception as e:
            print(f"    Error inserting SVG: {e}")
            try:
                shape = slide.Shapes.AddPicture(abs_path, True, False, 0, 0)
                original_width = shape.Width
                original_height = shape.Height
                original_ratio = original_width / original_height
                available_ratio = available_width / available_height
                if original_ratio > available_ratio:
                    new_width = available_width
                    new_height = available_width / original_ratio
                else:
                    new_height = available_height
                    new_width = available_height * original_ratio
                shape.Width = int(new_width)
                shape.Height = int(new_height)
                shape.Left = int((slide_width - new_width) / 2)
                shape.Top = int((slide_height - new_height) / 2)
                print(f"    Inserted (linked): {abs_path}")
            except Exception as e2:
                print(f"    Failed: {e2}")

    abs_output = os.path.abspath(output_path)
    presentation.SaveAs(abs_output)
    print(f"\nSaved: {abs_output}")

    return True


def get_next_pptx_path(output_dir, base_name="china_provinces_map"):
    """Get next available PPTX path with sequence number."""
    existing = [
        f
        for f in os.listdir(output_dir)
        if f.startswith(base_name) and f.endswith(".pptx")
    ]

    if not existing:
        return os.path.join(output_dir, f"{base_name}_01.pptx")

    # Extract sequence numbers from existing files
    seq_nums = []
    for f in existing:
        try:
            # Format: base_name_XX.pptx
            num = int(f.replace(base_name, "").replace(".pptx", "").replace("_", ""))
            seq_nums.append(num)
        except:
            pass

    if not seq_nums:
        return os.path.join(output_dir, f"{base_name}_01.pptx")

    next_seq = max(seq_nums) + 1
    return os.path.join(output_dir, f"{base_name}_{next_seq:02d}.pptx")


def main():
    os.makedirs(MAPS_DIR, exist_ok=True)

    print("=" * 60)
    print("Fetching China GeoJSON data...")
    print("=" * 60)

    china_data = fetch_geojson(CHINA_ADCODE, full=True)
    if not china_data:
        print("Failed to fetch China data!")
        return

    provinces = {}
    if "features" in china_data:
        for feature in china_data["features"]:
            props = feature.get("properties", {})
            name = props.get("name", "Unknown")
            adcode = str(props.get("adcode", ""))
            if name and adcode and not adcode.endswith("_JD"):
                provinces[name] = adcode

    print(f"\nFound {len(provinces)} provinces/regions")

    print("\n" + "=" * 60)
    print("Creating SVG maps...")
    print("=" * 60)

    all_province_data = {}
    svg_files = {}
    colors = generate_colors(len(provinces))

    for idx, (name, adcode) in enumerate(sorted(provinces.items())):
        print(f"\n[{idx + 1}/{len(provinces)}] Processing: {name} ({adcode})")

        safe_name = name.replace("/", "_").replace("\\", "_")
        cache_file = os.path.join(MAPS_DIR, f"{safe_name}_cache.json")

        if os.path.exists(cache_file):
            with open(cache_file, "r", encoding="utf-8") as f:
                province_data = json.load(f)
        else:
            province_data = fetch_geojson(adcode, full=False)
            if province_data:
                with open(cache_file, "w", encoding="utf-8") as f:
                    json.dump(province_data, f, ensure_ascii=False)

        if not province_data:
            print(f"  Skipping {name} - no data available")
            continue

        all_province_data[name] = province_data

        svg_path = os.path.join(MAPS_DIR, f"{safe_name}.svg")
        color = colors[idx % len(colors)]

        result = create_map_svg(province_data, svg_path, name, color)
        if result:
            svg_files[name] = result

    # China overview
    print("\n" + "=" * 60)
    print("Creating China overview...")
    print("=" * 60)

    china_svg = os.path.join(MAPS_DIR, "china_overview.svg")
    china_result = create_china_overview(all_province_data, china_svg)

    all_svg_files = {}
    if china_result:
        all_svg_files["中国行政区划全图"] = china_result
    all_svg_files.update(svg_files)

    print("\n" + "=" * 60)
    print("Creating PowerPoint...")
    print("=" * 60)

    pptx_path = get_next_pptx_path(SCRIPT_DIR, "china_provinces_map")
    create_pptx_with_com(all_svg_files, pptx_path)

    print("\n" + "=" * 60)
    print("COMPLETED!")
    print("=" * 60)


if __name__ == "__main__":
    main()
