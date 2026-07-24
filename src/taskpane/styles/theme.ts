import { createLightTheme, createDarkTheme, BrandVariants, Theme } from '@fluentui/react-components';

// NVIDIA-inspired green brand variants (#76B900)
export const nvidiaBrandVariants: BrandVariants = {
  10: "#020400",
  20: "#0F1C00",
  30: "#182E00",
  40: "#213D00",
  50: "#2A4E00",
  60: "#356100",
  70: "#417400",
  80: "#4F8800",
  90: "#5D9C00",
  100: "#6DB200",
  110: "#76B900", // Base brand color
  120: "#86C81A",
  130: "#95D730",
  140: "#A5E643",
  150: "#B4F556",
  160: "#C4FF6A"
};

export const customLightTheme: Theme = {
  ...createLightTheme(nvidiaBrandVariants),
  // Additional Copilot premium overrides can go here
};

export const customDarkTheme: Theme = {
  ...createDarkTheme(nvidiaBrandVariants),
  // Additional Copilot premium overrides can go here
};
