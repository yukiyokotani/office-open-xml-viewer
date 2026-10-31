import { nullRecord } from '../internal/own-key';

/** MS-OI29500 §2.1.1274: untransformed directional-light definitions.
 * Keep negative/greater-than-one RGB, per-light scale/offset and contribution
 * flags: clamping individual lights changes brightRoom and colored rigs.
 */
export interface RigLight {
  rgb: readonly number[];
  direction: readonly number[];
  scale: number;
  offset: number;
  specular: boolean;
  diffuse: boolean;
}
export interface RigDefinition {
  ambient: number;
  lights: readonly RigLight[];
}
const MODERN: Readonly<Record<string, RigDefinition>> = {
  threePt: {
    ambient: 0,
    lights: [
      {
        rgb: [1.141, 1.141, 1.141],
        direction: [-0.6515, -0.2693, -0.7093],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.5, 0.5, 0.5],
        direction: [0.8482, 0.2469, -0.4686],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.5634, -0.2812, 0.7769],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  balanced: {
    ambient: 0.13,
    lights: [
      {
        rgb: [1.05, 1.05, 1.05],
        direction: [0.5263, -0.4092, -0.7453],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [-0.9386, 0.3426, -0.041],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.5, 0.5, 0.5],
        direction: [0.0934, 0.763, 0.6396],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  soft: {
    ambient: 0.3,
    lights: [
      {
        rgb: [0.8, 0.8, 0.8],
        direction: [-0.6897, 0.2484, -0.6802],
        scale: 0.5,
        offset: 0.5,
        specular: true,
        diffuse: true,
      },
    ],
  },
  harsh: {
    ambient: 0.28,
    lights: [
      {
        rgb: [0.88, 0.88, 0.88],
        direction: [0.6689, -0.6755, -0.3104],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.88, 0.88, 0.88],
        direction: [-0.592, -0.7371, -0.326],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  flood: {
    ambient: 0.13,
    lights: [
      {
        rgb: [1.1, 1.1, 1.1],
        direction: [0.5685, -0.7651, -0.3022],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [1.1, 1.1, 1.1],
        direction: [-0.2366, -0.9595, -0.1531],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.55, 0.55, 0.55],
        direction: [-0.8982, 0.1386, -0.4171],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  contrasting: {
    ambient: 1.0,
    lights: [
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.0, -1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.0, 1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
    ],
  },
  morning: {
    ambient: 0,
    lights: [
      {
        rgb: [0.669, 0.648, 0.596],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.459, 0.454, 0.385],
        direction: [-0.2781, -0.4509, -0.8482],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.9, 0.86, 0.83],
        direction: [0.672, -0.6185, -0.4073],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.911, 0.846, 0.728],
        direction: [-0.1825, 0.968, 0.1722],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  sunrise: {
    ambient: 0,
    lights: [
      {
        rgb: [0.667, 0.63, 0.527],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.459, 0.459, 0.371],
        direction: [-0.2781, -0.4509, -0.8482],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.826, 0.712, 0.638],
        direction: [0.672, -0.6185, -0.4073],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [1.511, 1.319, 0.994],
        direction: [-0.1825, 0.968, 0.1722],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
    ],
  },
  sunset: {
    ambient: 0,
    lights: [
      {
        rgb: [0.672, 0.169, 0.169],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.459, 0.448, 0.327],
        direction: [0.0922, -0.3551, -0.9303],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.775, 0.612, 0.502],
        direction: [0.672, -0.6185, -0.4073],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.761, 0.69, 0.397],
        direction: [-0.424, 0.8891, 0.1722],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
    ],
  },
  chilly: {
    ambient: 0.11,
    lights: [
      {
        rgb: [0.31, 0.32, 0.32],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.45, 0.45, 0.45],
        direction: [-0.3539, -0.1505, -0.9231],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [1.03, 1.02, 1.15],
        direction: [0.672, -0.6185, -0.4073],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.41, 0.45, 0.48],
        direction: [-0.5781, 0.7976, 0.1722],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  freezing: {
    ambient: 0,
    lights: [
      {
        rgb: [0.53, 0.567, 0.661],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.37, 0.461, 0.461],
        direction: [-0.2781, -0.4509, -0.8482],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.649, 0.638, 0.904],
        direction: [0.672, -0.6185, -0.4073],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.971, 1.19, 1.363],
        direction: [-0.1825, 0.968, 0.1722],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  flat: {
    ambient: 1.0,
    lights: [
      {
        rgb: [0.821, 0.821, 0.821],
        direction: [-0.9546, -0.1619, -0.2502],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
      {
        rgb: [2.072, 2.54, 2.91],
        direction: [0.0009, 0.8605, 0.5095],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
      {
        rgb: [3.843, 3.843, 3.843],
        direction: [0.6574, -0.7316, -0.1806],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
    ],
  },
  twoPt: {
    ambient: 0.25,
    lights: [
      {
        rgb: [0.84, 0.84, 0.84],
        direction: [0.5266, -0.4089, -0.7454],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.3, 0.3, 0.3],
        direction: [-0.8983, 0.2365, -0.3704],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  glow: {
    ambient: 1.0,
    lights: [
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.0, -1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
      {
        rgb: [0.7, 0.7, 0.7],
        direction: [0.0, 1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: true,
      },
    ],
  },
  brightRoom: {
    ambient: 1.5,
    lights: [
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.0, -1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [1.0, 1.0, 1.0],
        direction: [0.8227, -0.1882, -0.5364],
        scale: 1.0,
        offset: 0.0,
        specular: true,
        diffuse: false,
      },
      {
        rgb: [-0.5, -0.5, -0.5],
        direction: [0.0, 0.0, -1.0],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
      {
        rgb: [0.5, 0.5, 0.5],
        direction: [0.0, 1.0, 0.0],
        scale: 1.0,
        offset: 0.0,
        specular: false,
        diffuse: true,
      },
    ],
  },
};

/** MS-OI29500 §2.1.1324: legacy key is specular+diffuse, fill diffuse only. */
const LEGACY: Record<string, RigDefinition> = nullRecord();
for (const [family, key, fill, ambient, fillScale, fillOffset] of [
  ['Flat', 0.58, 0.58, 0.305, 0.5, 0.5],
  ['Normal', 0.671, 0.366, 0.153, 0.5, 0.5],
  ['Harsh', 0.793, 0.214, 0.061, 1, 0],
] as const) {
  for (let i = 1; i <= 4; i++) {
    const direction = i === 1 ? [0, 0, -0.2] : [i - 3, -1, -0.2];
    const fillDirection = i === 1 ? [0, 0, -0.2] : [0, 1, -0.2];
    LEGACY[`legacy${family}${i}`] = {
      ambient,
      lights: [
        {
          rgb: [key, key, key],
          direction,
          scale: 1,
          offset: 0,
          specular: true,
          diffuse: true,
        },
        {
          rgb: [fill, fill, fill],
          direction: fillDirection,
          scale: fillScale,
          offset: fillOffset,
          specular: false,
          diffuse: true,
        },
      ],
    };
  }
}
export function rigDefinition(name: string): RigDefinition {
  if (Object.hasOwn(MODERN, name)) return MODERN[name];
  return Object.hasOwn(LEGACY, name) ? LEGACY[name] : MODERN.threePt;
}
