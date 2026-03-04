/**
 * Training Media — single course entity-card with displayMode-aware fullscreen.
 *
 * Inline : single entity card with video player, metadata, badges,
 *               description, and action buttons.
 * Fullscreen:   embedded video player, metadata, instructor info, sidebar actions.
 *               Driven by `hostContext.displayMode`.
 */
import {
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowMaximizeRegular,
  BookmarkRegular,
  CertificateRegular,
  ClockRegular,
  DocumentArrowDownRegular,
  PeopleRegular,
  ShareRegular,
} from "@fluentui/react-icons";
import { StrictMode, useCallback, useEffect } from "react";
import { createRoot } from "react-dom/client";
import { FluentWrapper } from "../shared/FluentWrapper";
import { useToolData } from "../shared/useToolData";
import type { TrainingMediaData, Course, CourseModule } from "../../mock-data/training-media";

// ─── Icon for module type ────────────────────────────────────────────────────
import {
  VideoRegular,
  BookOpenRegular,
  BeakerRegular,
} from "@fluentui/react-icons";

// ─── Inline Card Styles ──────────────────────────────────────────────────────

const useCardStyles = makeStyles({
  // Outer entity card
  card: {
    backgroundColor: tokens.colorStrokeFocus1,
    paddingInlineStart: tokens.spacingHorizontalXXL,
    paddingInlineEnd: tokens.spacingHorizontalXXL,
    paddingBlockStart: tokens.spacingVerticalXXL,
    paddingBlockEnd: tokens.spacingVerticalXXL,
    borderRadius: "24px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    rowGap: tokens.spacingVerticalL,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    width: "100%",
    boxSizing: "border-box",
  },

  // Title row — title + expand button
  titleRow: {
    columnGap: tokens.spacingHorizontalS,
    display: "flex",
    alignItems: "center",
    alignSelf: "stretch",
  },
  titleText: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    flexGrow: 1,
    flexShrink: 1,
    flexBasis: "0%",
    minWidth: 0,
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
    display: "-webkit-box",
    WebkitBoxOrient: "vertical",
    ...shorthands.overflow("hidden"),
    textOverflow: "ellipsis",
  },
  expandBtn: {
    backgroundColor: tokens.colorStrokeFocus1,
    paddingInlineStart: tokens.spacingHorizontalS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingBlockEnd: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusLarge,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    border: "none",
    cursor: "pointer",
    color: tokens.colorNeutralForeground1Static,
    "&:hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },

  // Metadata row
  metadataRow: {
    columnGap: tokens.spacingHorizontalXXS,
    display: "flex",
    height: "20px",
    alignItems: "center",
    alignSelf: "stretch",
  },
  metadataItems: {
    columnGap: tokens.spacingHorizontalS,
    display: "flex",
    alignItems: "center",
  },
  metaText: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase200,
    color: "#7a7a7a",
    display: "-webkit-box",
    WebkitBoxOrient: "vertical",
    ...shorthands.overflow("hidden"),
    textOverflow: "ellipsis",
  },
  metaDot: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase200,
    color: "#999999",
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    alignItems: "center",
  },

  // Thumbnail + badges wrapper
  thumbnailWrap: {
    rowGap: tokens.spacingVerticalS,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    alignSelf: "stretch",
  },
  thumbnailContainer: {
    borderRadius: "16px",
    alignSelf: "stretch",
    position: "relative",
    aspectRatio: "16/9",
    ...shorthands.overflow("hidden"),
  },
  thumbnailImg: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
    display: "block",
  },
  playOverlay: {
    position: "absolute",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    backgroundColor: "rgba(0,0,0,0.55)",
    paddingInlineStart: "14px",
    paddingInlineEnd: "14px",
    paddingBlockStart: "14px",
    paddingBlockEnd: "14px",
    borderRadius: "999px",
    display: "inline-flex",
    alignItems: "center",
    backdropFilter: "blur(6px)",
    color: "#fff",
    cursor: "pointer",
  },
  videoIframe: {
    width: "100%",
    height: "100%",
    border: "none",
    display: "block",
  },

  // Badge chips
  badgeRow: {
    display: "flex",
    alignItems: "flex-start",
    alignSelf: "stretch",
    columnGap: tokens.spacingHorizontalXS,
    flexWrap: "wrap",
  },
  badgeChip: {
    backgroundColor: tokens.colorStrokeFocus1,
    paddingInlineStart: tokens.spacingHorizontalS,
    paddingInlineEnd: tokens.spacingHorizontalS,
    paddingBlockStart: tokens.spacingVerticalXXS,
    paddingBlockEnd: tokens.spacingVerticalXXS,
    borderRadius: tokens.borderRadiusLarge,
    columnGap: tokens.spacingHorizontalXS,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
  },
  badgeIcon: {
    width: "20px",
    height: "20px",
    color: tokens.colorNeutralForeground2Link,
  },
  badgeText: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase200,
    color: tokens.colorNeutralForeground2Link,
  },

  // Description
  description: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1Static,
    alignSelf: "stretch",
    display: "-webkit-box",
    WebkitBoxOrient: "vertical",
    WebkitLineClamp: 3,
    ...shorthands.overflow("hidden"),
    textOverflow: "ellipsis",
  },

  // Action buttons
  actionsRow: {
    columnGap: tokens.spacingHorizontalS,
    display: "flex",
    alignItems: "center",
  },
  enrollBtn: {
    backgroundColor: "#464feb",
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingBlockEnd: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusXLarge,
    columnGap: tokens.spacingHorizontalSNudge,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    border: "none",
    cursor: "pointer",
  },
  enrollBtnText: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorStrokeFocus1,
  },
  openAppBtn: {
    backgroundColor: "transparent",
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingBlockEnd: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusXLarge,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    cursor: "pointer",
  },
  openAppBtnText: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
  },
});

// ─── Fullscreen Styles ───────────────────────────────────────────────────────

const useFullscreenStyles = makeStyles({
  canvas: {
    backgroundColor: "#feffff",
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    width: "100%",
    minHeight: "100vh",
    rowGap: tokens.spacingVerticalXXL,
  },

  // Top section: video + course modules sidebar
  topSection: {
    display: "flex",
    alignItems: "flex-start",
    width: "100%",
  },

  // Video column — takes ~75% of remaining space
  videoCol: {
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingBlockEnd: tokens.spacingVerticalL,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    flexGrow: 1,
    flexShrink: 1,
    flexBasis: 0,
    minWidth: 0,
    rowGap: tokens.spacingVerticalL,
    width: "calc(100% - 320px)",
  },
  videoContainer: {
    borderRadius: "16px",
    alignSelf: "stretch",
    ...shorthands.overflow("hidden"),
    position: "relative",
    aspectRatio: "16/9",
  },
  videoIframe: {
    width: "100%",
    height: "100%",
    border: "none",
    display: "block",
  },

  // Metadata row
  metaRow: {
    columnGap: tokens.spacingHorizontalS,
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
  },
  metaText: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase300,
    color: "#7a7a7a",
  },
  metaDotWrap: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
  },
  metaDot: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase300,
    color: "#999999",
  },

  // Instructor row
  instructorRow: {
    paddingInlineStart: tokens.spacingHorizontalXS,
    columnGap: tokens.spacingHorizontalM,
    display: "flex",
    alignItems: "center",
  },
  avatar: {
    borderRadius: tokens.borderRadiusCircular,
    width: "32px",
    height: "32px",
    backgroundColor: "#464feb",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#fff",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    flexShrink: 0,
  },
  instructorInfo: {
    rowGap: tokens.spacingVerticalXXS,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
  },
  instructorName: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
  },
  badge: {
    backgroundColor: tokens.colorBrandBackgroundInvertedHover,
    paddingInlineStart: tokens.spacingHorizontalXXS,
    paddingInlineEnd: tokens.spacingHorizontalXXS,
    borderRadius: tokens.borderRadiusMedium,
    display: "inline-flex",
    height: "16px",
    justifyContent: "center",
    alignItems: "center",
  },
  badgeText: {
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase100,
    color: tokens.colorCompoundBrandStrokeHover,
    textAlign: "center",
    paddingInlineStart: tokens.spacingHorizontalXXS,
    paddingInlineEnd: tokens.spacingHorizontalXXS,
  },

  // ── Course modules sidebar ──
  sidebar: {
    width: "320px",
    minWidth: "320px",
    maxWidth: "320px",
    flexShrink: 0,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    alignSelf: "stretch",
    ...shorthands.overflow("hidden"),
  },
  sidebarHeader: {
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalL,
    paddingBlockEnd: tokens.spacingVerticalM,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  sidebarTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
  },
  moduleList: {
    display: "flex",
    flexDirection: "column",
    overflowY: "auto",
  },
  moduleItem: {
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingBlockEnd: tokens.spacingVerticalS,
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalM,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "pointer",
    "&:hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  moduleIcon: {
    width: "20px",
    height: "20px",
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  moduleInfo: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    minWidth: 0,
  },
  moduleTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightMedium,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1Static,
    whiteSpace: "nowrap",
    ...shorthands.overflow("hidden"),
    textOverflow: "ellipsis",
  },
  moduleMeta: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase200,
    color: "#7a7a7a",
  },

  // ── Bottom content area ──
  bottomSection: {
    paddingInlineStart: tokens.spacingHorizontalXXL,
    paddingInlineEnd: tokens.spacingHorizontalXXL,
    paddingBlockEnd: tokens.spacingVerticalXXL,
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXL,
  },

  // Rich description
  descriptionBlock: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    maxWidth: "800px",
  },
  descriptionText: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
  },
  readMoreBtn: {
    backgroundColor: "transparent",
    border: "none",
    cursor: "pointer",
    padding: 0,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase300,
    color: "#464feb",
    alignSelf: "flex-start",
    "&:hover": { textDecorationLine: "underline" },
  },

  // Enroll row
  enrollBtn: {
    backgroundColor: "#464feb",
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalS,
    paddingBlockEnd: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusXLarge,
    display: "inline-flex",
    justifyContent: "center",
    alignItems: "center",
    border: "none",
    cursor: "pointer",
    alignSelf: "flex-start",
  },
  enrollBtnText: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorStrokeFocus1,
  },

  // ── Related courses ──
  relatedSection: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalM,
  },
  relatedHeading: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1Static,
  },
  relatedGrid: {
    display: "flex",
    columnGap: tokens.spacingHorizontalL,
    overflowX: "auto",
    scrollbarWidth: "none",
    "&::-webkit-scrollbar": { display: "none" },
  },
  relatedCard: {
    width: "200px",
    minWidth: "200px",
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    cursor: "pointer",
    borderRadius: tokens.borderRadiusLarge,
    "&:hover": { opacity: "0.85" },
  },
  relatedThumb: {
    width: "100%",
    height: "112px",
    borderRadius: tokens.borderRadiusLarge,
    display: "block",
    border: "none",
    pointerEvents: "none" as const,
  },
  relatedTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1Static,
    display: "-webkit-box",
    WebkitBoxOrient: "vertical",
    WebkitLineClamp: 2,
    ...shorthands.overflow("hidden"),
    textOverflow: "ellipsis",
  },
  relatedInstructor: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase200,
    color: "#7a7a7a",
  },

  // ── Sidebar action buttons (Save, Share, Download) ──
  actionsSection: {
    paddingInlineStart: tokens.spacingHorizontalL,
    paddingInlineEnd: tokens.spacingHorizontalL,
    paddingBlockStart: tokens.spacingVerticalM,
    paddingBlockEnd: tokens.spacingVerticalM,
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  actionsHeading: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1Static,
    letterSpacing: "-0.15px",
    paddingBlockEnd: tokens.spacingVerticalXS,
  },
  actionBtn: {
    borderRadius: "10px",
    display: "flex",
    alignItems: "center",
    alignSelf: "stretch",
    columnGap: tokens.spacingHorizontalS,
    backgroundColor: "transparent",
    border: "none",
    cursor: "pointer",
    padding: tokens.spacingHorizontalS,
    "&:hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  actionIcon: {
    width: "16px",
    height: "16px",
    color: tokens.colorNeutralForeground2Link,
  },
  actionText: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightMedium,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground2Link,
    letterSpacing: "-0.15px",
  },
});

// ─── Inline Entity Card ──────────────────────────────────────────────────────

function EntityCard({
  course,
  onExpand,
}: {
  course: Course;
  onExpand: () => void;
}) {
  const s = useCardStyles();

  return (
    <div className={s.card}>
      {/* Title + metadata + expand */}
      <div className={s.titleRow}>
        <div className={s.titleText}>
          <div className={s.title}>{course.title}</div>
          <div className={s.metadataRow}>
            <div className={s.metadataItems}>
              <div className={s.metaText}>{course.duration}</div>
              <div className={s.metaDot}>•</div>
              <div className={s.metaText}>{course.lessonsCount} lessons</div>
              <div className={s.metaDot}>•</div>
              <div className={s.metaText}>{course.level}</div>
            </div>
          </div>
        </div>
        <button className={s.expandBtn} onClick={onExpand} aria-label="Expand">
          <ArrowMaximizeRegular style={{ width: 24, height: 24 }} />
        </button>
      </div>

      {/* Video player + badges */}
      <div className={s.thumbnailWrap}>
        <div className={s.thumbnailContainer}>
          <iframe
            className={s.videoIframe}
            src={course.videoUrl}
            title={course.title}
            allow="autoplay; encrypted-media; picture-in-picture"
          />
        </div>

        <div className={s.badgeRow}>
          <div className={s.badgeChip}>
            <PeopleRegular className={s.badgeIcon} />
            <span className={s.badgeText}>{course.enrolledCount} enrolled</span>
          </div>
          {course.hasCertificate && (
            <div className={s.badgeChip}>
              <CertificateRegular className={s.badgeIcon} />
              <span className={s.badgeText}>Certificate</span>
            </div>
          )}
          <div className={s.badgeChip}>
            <ClockRegular className={s.badgeIcon} />
            <span className={s.badgeText}>{course.pacing}</span>
          </div>
        </div>
      </div>

      {/* Description */}
      <div className={s.description}>{course.description}</div>

      {/* Actions */}
      <div className={s.actionsRow}>
        <button className={s.enrollBtn}>
          <span className={s.enrollBtnText}>Enroll for free</span>
        </button>
      </div>
    </div>
  );
}

// ─── Fullscreen View ─────────────────────────────────────────────────────────

function FullscreenView({ course, data }: { course: Course; data: TrainingMediaData }) {
  const s = useFullscreenStyles();
  const initials = course.instructor
    .split(" ")
    .map((w) => w[0])
    .join("");

  const relatedCourses = data.courses.filter((c) => c.id !== course.id);

  const moduleIcon = (type: string) => {
    switch (type.toLowerCase()) {
      case "video":
        return <VideoRegular className={s.moduleIcon} />;
      case "lab":
        return <BeakerRegular className={s.moduleIcon} />;
      default:
        return <BookOpenRegular className={s.moduleIcon} />;
    }
  };

  return (
    <div className={s.canvas}>
      {/* Top section: video + course modules sidebar */}
      <div className={s.topSection}>
        {/* Video column */}
        <div className={s.videoCol}>
          {/* Video */}
          <div className={s.videoContainer}>
            <iframe
              className={s.videoIframe}
              src={course.videoUrl}
              title={course.title}
              allow="autoplay; encrypted-media; picture-in-picture"
            />
          </div>

          {/* Metadata */}
          <div className={s.metaRow}>
            <div className={s.metaText}>{course.duration}</div>
            <div className={s.metaDotWrap}>
              <div className={s.metaDot}>•</div>
            </div>
            <div className={s.metaText}>{course.lessonsCount} lessons</div>
            <div className={s.metaDotWrap}>
              <div className={s.metaDot}>•</div>
            </div>
            <div className={s.metaText}>{course.level}</div>
            <div className={s.metaDotWrap}>
              <div className={s.metaDot}>•</div>
            </div>
            <div className={s.metaText}>{course.year}</div>
          </div>

          {/* Instructor */}
          <div className={s.instructorRow}>
            <div className={s.avatar}>{initials}</div>
            <div className={s.instructorInfo}>
              <div className={s.instructorName}>
                Instructor: {course.instructor}
              </div>
              <div className={s.badge}>
                <span className={s.badgeText}>{course.instructorBadge}</span>
              </div>
            </div>
          </div>

          {/* Enroll */}
          <button className={s.enrollBtn}>
            <span className={s.enrollBtnText}>Enroll for free</span>
          </button>
        </div>

        {/* Course modules sidebar */}
        <div className={s.sidebar}>
          <div className={s.sidebarHeader}>
            <div className={s.sidebarTitle}>Course modules</div>
          </div>
          <div className={s.moduleList}>
            {(course.modules ?? []).map((mod: CourseModule, i: number) => (
              <div className={s.moduleItem} key={i}>
                {moduleIcon(mod.type)}
                <div className={s.moduleInfo}>
                  <div className={s.moduleTitle}>{mod.title}</div>
                  <div className={s.moduleMeta}>
                    {mod.type} · {mod.duration}
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Actions at bottom of sidebar */}
          <div className={s.actionsSection}>
            <div className={s.actionsHeading}>Actions</div>
            <button className={s.actionBtn}>
              <BookmarkRegular className={s.actionIcon} />
              <span className={s.actionText}>Save for later</span>
            </button>
            <button className={s.actionBtn}>
              <ShareRegular className={s.actionIcon} />
              <span className={s.actionText}>Share with team</span>
            </button>
            <button className={s.actionBtn}>
              <DocumentArrowDownRegular className={s.actionIcon} />
              <span className={s.actionText}>Download transcript</span>
            </button>
          </div>
        </div>
      </div>

      {/* Bottom content area */}
      <div className={s.bottomSection}>
        {/* Rich description */}
        <div className={s.descriptionBlock}>
          <div className={s.descriptionText}>
            {course.richDescription || course.description}
          </div>
        </div>

        {/* Related courses */}
        {relatedCourses.length > 0 && (
          <div className={s.relatedSection}>
            <div className={s.relatedHeading}>Related courses</div>
            <div className={s.relatedGrid}>
              {relatedCourses.map((rc) => (
                <div className={s.relatedCard} key={rc.id}>
                  <iframe
                    className={s.relatedThumb}
                    src={rc.videoUrl}
                    title={rc.title}
                    tabIndex={-1}
                  />
                  <div className={s.relatedTitle}>{rc.title}</div>
                  <div className={s.relatedInstructor}>{rc.instructor}</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Root ────────────────────────────────────────────────────────────────────

function TrainingMediaApp() {
  const { data, app, hostContext } = useToolData<TrainingMediaData>("training-media");

  // Determine display mode from host context
  const displayMode = hostContext?.displayMode ?? "inline";
  const isFullscreen = displayMode === "fullscreen";

  // Pick the first course to display
  const course = data?.courses[0] ?? null;

  // Hide overflow for inline mode to prevent scrollbar, allow scroll in fullscreen
  useEffect(() => {
    document.documentElement.style.overflow = isFullscreen ? "auto" : "hidden";
    document.body.style.overflow = isFullscreen ? "auto" : "hidden";
    return () => {
      document.documentElement.style.overflow = "";
      document.body.style.overflow = "";
    };
  }, [isFullscreen]);

  const handleExpand = useCallback(async () => {
    if (!app) return;
    const ctx = app.getHostContext();
    if (ctx?.availableDisplayModes?.includes("fullscreen")) {
      await app.requestDisplayMode({ mode: "fullscreen" });
    }
  }, [app]);

  if (!course) {
    return (
      <FluentWrapper app={app} hostContext={hostContext}>
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            padding: "40px 0",
            color: tokens.colorNeutralForeground3,
          }}
        >
          Loading course…
        </div>
      </FluentWrapper>
    );
  }

  return (
    <FluentWrapper app={app} hostContext={hostContext}>
      {isFullscreen ? (
        <FullscreenView course={course} data={data!} />
      ) : (
        <EntityCard course={course} onExpand={handleExpand} />
      )}
    </FluentWrapper>
  );
}

// ─── Mount ────────────────────────────────────────────────────────────────────
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <TrainingMediaApp />
  </StrictMode>,
);
