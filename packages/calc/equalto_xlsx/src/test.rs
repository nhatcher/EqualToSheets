use super::colors::*;

#[test]
fn test_known_colors() {
    let color1 = get_themed_color(0, -0.05);
    assert_eq!(color1, "#F2F2F2");

    let color2 = get_themed_color(5, -0.25);
    // Excel returns "#C65911" (rounding error)
    assert_eq!(color2, "#C55911");

    let color3 = get_themed_color(4, 0.6);
    // Excel returns "#b4c6e7" (rounding error)
    assert_eq!(color3, "#B5C8E8");
}

#[test]
fn test_rgb_hex() {
    struct ColorTest {
        hex: String,
        rgb: [i32; 3],
        hsl: [i32; 3],
    }
    let color_tests = [
        ColorTest {
            hex: "#FFFFFF".to_string(),
            rgb: [255, 255, 255],
            hsl: [0, 0, 100],
        },
        ColorTest {
            hex: "#000000".to_string(),
            rgb: [0, 0, 0],
            hsl: [0, 0, 0],
        },
        ColorTest {
            hex: "#44546A".to_string(),
            rgb: [68, 84, 106],
            hsl: [215, 22, 34],
        },
        ColorTest {
            hex: "#E7E6E6".to_string(),
            rgb: [231, 230, 230],
            hsl: [0, 2, 90],
        },
        ColorTest {
            hex: "#4472C4".to_string(),
            rgb: [68, 114, 196],
            hsl: [218, 52, 52],
        },
        ColorTest {
            hex: "#ED7D31".to_string(),
            rgb: [237, 125, 49],
            hsl: [24, 84, 56],
        },
        ColorTest {
            hex: "#A5A5A5".to_string(),
            rgb: [165, 165, 165],
            hsl: [0, 0, 65],
        },
        ColorTest {
            hex: "#FFC000".to_string(),
            rgb: [255, 192, 0],
            hsl: [45, 100, 50],
        },
        ColorTest {
            hex: "#5B9BD5".to_string(),
            rgb: [91, 155, 213],
            hsl: [209, 59, 60],
        },
        ColorTest {
            hex: "#70AD47".to_string(),
            rgb: [112, 173, 71],
            hsl: [96, 42, 48],
        },
        ColorTest {
            hex: "#0563C1".to_string(),
            rgb: [5, 99, 193],
            hsl: [210, 95, 39],
        },
        ColorTest {
            hex: "#954F72".to_string(),
            rgb: [149, 79, 114],
            hsl: [330, 31, 45],
        },
    ];
    for color in color_tests.iter() {
        let rgb = color.rgb;
        let hsl = color.hsl;
        assert_eq!(rgb, hex_to_rgb(&color.hex));
        assert_eq!(hsl, rgb_to_hsl(rgb));
        assert_eq!(rgb_to_hex(rgb), color.hex);
        // The round trip has rounding errors
        // FIXME: We could also hardcode the hsl21 in the testcase
        let rgb2 = hsl_to_rgb(hsl);
        let diff = (rgb2[0] - rgb[0]).abs() + (rgb2[1] - rgb[1]).abs() + (rgb2[2] - rgb[2]).abs();
        assert!(diff < 4);
    }
}
