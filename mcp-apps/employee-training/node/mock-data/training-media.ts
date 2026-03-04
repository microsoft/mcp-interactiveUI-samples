/**
 * Mock data for Training Media tool.
 * 5 AI/ML courses with Microsoft Learn Video embeds.
 */

export interface CourseModule {
  title: string;
  type: string;
  duration: string;
}

export interface Course {
  id: string;
  title: string;
  duration: string;
  lessonsCount: number;
  level: string;
  year: string;
  description: string;
  richDescription: string;
  enrolledCount: string;
  hasCertificate: boolean;
  pacing: string;
  instructor: string;
  instructorBadge: string;
  videoUrl: string;
  modules: CourseModule[];
}

export interface TrainingMediaData {
  heading: string;
  subtitle: string;
  courses: Course[];
}

export const trainingMediaData: TrainingMediaData = {
  heading: "AI & Microsoft Learn",
  subtitle: "Select a course to start learning",
  courses: [
    {
      id: "course-001",
      title: "All About Azure AI Agents",
      duration: "27m",
      lessonsCount: 7,
      level: "Intermediate",
      year: "2025",
      description:
        "Explore Azure AI Agents — from Fabric data agents to connected agents and computer use agents (CUA). Learn what problems they solve and when to use them.",
      richDescription:
        "Explore Azure AI Agents — from Fabric data agents to connected agents and computer use agents (CUA). In this episode of The AI Show, Marco Casalaina joins Seth Juarez to break down the Azure AI Agent Service, showing how agents are solving real-world problems. You'll see demos of Fabric data agents, connected agents that compose multiple capabilities, and the cutting-edge computer use agent (CUA) that interacts with UIs directly.",
      enrolledCount: "15,230",
      hasCertificate: true,
      pacing: "Self-paced",
      instructor: "Seth Juarez",
      instructorBadge: "AI Show",
      videoUrl: "https://learn-video.azurefd.net/vod/player?show=ai-show&ep=all-about-azure-ai-agents",
      modules: [
        { title: "Introduction to Azure AI Agents", type: "Video", duration: "1m" },
        { title: "Problems Agents Are Solving", type: "Video", duration: "3m" },
        { title: "Fabric Data Agent", type: "Video", duration: "5m" },
        { title: "Connected Agents", type: "Video", duration: "5m" },
        { title: "Computer Use Agent (CUA)", type: "Video", duration: "6m" },
        { title: "When to Use CUA", type: "Reading", duration: "4m" },
        { title: "Bonus Content", type: "Video", duration: "3m" },
      ],
    },
    {
      id: "course-002",
      title: "Semantic Kernel's Agent Framework",
      duration: "25m",
      lessonsCount: 5,
      level: "Intermediate",
      year: "2025",
      description:
        "Discover Semantic Kernel's Agent Framework — build multi-agent systems, compose AI capabilities, and orchestrate agents with the Aspire dashboard.",
      richDescription:
        "Discover Semantic Kernel's Agent Framework — build multi-agent systems, compose AI capabilities, and orchestrate agents. Shawn Henry joins Seth Juarez to discuss how Semantic Kernel enables agentic behavior, multi-agent collaboration, and interop with other frameworks. You'll see a live demo of agents working together and learn how the Aspire dashboard helps you trace and debug agent workflows in production.",
      enrolledCount: "9,812",
      hasCertificate: true,
      pacing: "Self-paced",
      instructor: "Seth Juarez",
      instructorBadge: "AI Show",
      videoUrl: "https://learn-video.azurefd.net/vod/player?show=ai-show&ep=semantic-kernels-agent-framework",
      modules: [
        { title: "What is Semantic Kernel?", type: "Video", duration: "3m" },
        { title: "Enabling Agentic Behavior", type: "Video", duration: "5m" },
        { title: "Multi-Agent Systems", type: "Video", duration: "3m" },
        { title: "Live Demo", type: "Lab", duration: "9m" },
        { title: "Aspire Dashboard & Interop", type: "Video", duration: "5m" },
      ],
    },
    {
      id: "course-003",
      title: "Best Practices for Agentic Apps with Azure AI Foundry",
      duration: "35m",
      lessonsCount: 4,
      level: "Advanced",
      year: "2025",
      description:
        "Go behind the scenes of the Build keynote to learn best practices for building agentic applications using Azure AI Foundry.",
      richDescription:
        "Go behind the scenes of the Build keynote to learn best practices for building agentic applications using Azure AI Foundry. Amanda Foster joins Seth Juarez to break down how the keynote agents were created, share architectural patterns, and reveal the engineering decisions that went into building production-grade agentic apps. Whether you're just getting started or scaling your AI solutions, you'll walk away with actionable guidance.",
      enrolledCount: "12,543",
      hasCertificate: true,
      pacing: "Self-paced",
      instructor: "Seth Juarez",
      instructorBadge: "AI Show",
      videoUrl: "https://learn-video.azurefd.net/vod/player?show=ai-show&ep=best-practices-for-building-agentic-apps-with-azure-ai-foundry",
      modules: [
        { title: "Introduction", type: "Video", duration: "1m" },
        { title: "Recap of Build Keynote", type: "Video", duration: "22m" },
        { title: "Insights into the Agents", type: "Video", duration: "4m" },
        { title: "Advice for Building Agentic Apps", type: "Video", duration: "8m" },
      ],
    },
    {
      id: "course-004",
      title: "Agentic Retrieval in Azure AI Search",
      duration: "15m",
      lessonsCount: 4,
      level: "Intermediate",
      year: "2025",
      description:
        "Say goodbye to one-shot RAG! Learn about agentic retrieval in Azure AI Search — a smart, multiturn query engine that plans and executes its own retrieval strategy.",
      richDescription:
        "Say goodbye to one-shot RAG! Microsoft just launched agentic retrieval in Azure AI Search — a smart, multiturn query engine that plans and executes its own retrieval strategy to tackle complex questions. Pablo Castro joins Seth Juarez to explain how agentic retrieval works, demonstrate it live, and show evaluation results that prove its effectiveness over traditional approaches.",
      enrolledCount: "28,910",
      hasCertificate: false,
      pacing: "Self-paced",
      instructor: "Seth Juarez",
      instructorBadge: "AI Show",
      videoUrl: "https://learn-video.azurefd.net/vod/player?show=ai-show&ep=agentic-retrieval-in-azure-ai-search",
      modules: [
        { title: "Introduction & What's New", type: "Video", duration: "1m" },
        { title: "Agentic Retrieval Explained", type: "Video", duration: "4m" },
        { title: "Live Demo", type: "Lab", duration: "6m" },
        { title: "Evaluation Results & Learn More", type: "Reading", duration: "4m" },
      ],
    },
    {
      id: "course-005",
      title: "Build Your Own Copilot with Azure AI Studio",
      duration: "19m",
      lessonsCount: 6,
      level: "Beginner",
      year: "2024",
      description:
        "Learn to create generative AI apps like an enterprise chat copilot using prebuilt and customizable models from Azure OpenAI Service.",
      richDescription:
        "Learn to create generative AI apps like an enterprise chat copilot using prebuilt and customizable models from Azure OpenAI Service. This episode walks you through Azure AI Studio end-to-end — from getting started and deploying your first model to adding data sources, building indexes, running manual evaluations, and scaling your proof of concept to production with prompt flow.",
      enrolledCount: "6,421",
      hasCertificate: false,
      pacing: "Self-paced",
      instructor: "Seth Juarez",
      instructorBadge: "AI Show",
      videoUrl: "https://learn-video.azurefd.net/vod/player?show=ai-show&ep=build-your-own-copilot-with-azure-ai-studio-part-1",
      modules: [
        { title: "Overview of Azure AI Studio", type: "Video", duration: "2m" },
        { title: "Getting Started", type: "Video", duration: "3m" },
        { title: "Adding a Data Source", type: "Reading", duration: "3m" },
        { title: "Building an Index", type: "Video", duration: "3m" },
        { title: "Manual Evaluations", type: "Video", duration: "3m" },
        { title: "Part 1 Summary & Next Steps", type: "Video", duration: "2m" },
      ],
    },
  ],
};
