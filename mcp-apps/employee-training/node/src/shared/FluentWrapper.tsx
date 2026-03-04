/**
 * FluentWrapper — Shared wrapper providing Fluent UI v9 theming and host integration.
 *
 * - Wraps children in FluentProvider with light/dark theme based on host context
 * - Applies safe area insets from the MCP host
 * - Mobile-first responsive container
 */
import type { App, McpUiHostContext } from "@modelcontextprotocol/ext-apps";
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { ReactNode } from "react";

const useStyles = makeStyles({
  root: {
    width: "100%",
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
  },
});

interface FluentWrapperProps {
  app: App | null;
  hostContext?: McpUiHostContext;
  children: ReactNode;
}

export function FluentWrapper({ hostContext, children }: FluentWrapperProps) {
  const isDark = hostContext?.theme === "dark";
  const theme = isDark ? webDarkTheme : webLightTheme;
  const styles = useStyles();

  const safeArea = hostContext?.safeAreaInsets;

  return (
    <FluentProvider theme={theme} style={{ background: "transparent" }}>
      <div
        className={styles.root}
        style={{
          paddingTop: safeArea?.top ? `${safeArea.top}px` : undefined,
          paddingRight: safeArea?.right ? `${safeArea.right}px` : undefined,
          paddingBottom: safeArea?.bottom ? `${safeArea.bottom}px` : undefined,
          paddingLeft: safeArea?.left ? `${safeArea.left}px` : undefined,
        }}
      >
        {children}
      </div>
    </FluentProvider>
  );
}
