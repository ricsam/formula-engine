import { createRootRoute, Outlet } from "@tanstack/react-router";
import { TanStackRouterDevtools } from "@tanstack/react-router-devtools";
import { useEffect, useState } from "react";
import { Sidebar } from "../components/Sidebar";
import { MobileMenuToggle } from "../components/MobileMenuToggle";
import { PanelLeftClose, PanelLeftOpen } from "lucide-react";
import "../index.css";

export const Route = createRootRoute({
  component: () => {
    const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
    const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);

    const toggleMobileMenu = () => {
      setIsMobileMenuOpen(!isMobileMenuOpen);
    };

    const closeMobileMenu = () => {
      setIsMobileMenuOpen(false);
    };

    const toggleSidebar = () => {
      setIsSidebarCollapsed(!isSidebarCollapsed);
    };

    // useEffect(() => {
    //   const root = window.document.documentElement;

    //   root.classList.remove("light", "dark");

    //   const systemTheme = window.matchMedia("(prefers-color-scheme: dark)")
    //     .matches
    //     ? "dark"
    //     : "light";

    //   root.classList.add(systemTheme);
    // }, []);

    return (
      <div className="flex h-full w-full">
        {/* Mobile Menu Toggle */}
        <MobileMenuToggle
          isOpen={isMobileMenuOpen}
          onToggle={toggleMobileMenu}
        />

        {/* Desktop Expand Button (when sidebar is collapsed) */}
        {isSidebarCollapsed && (
          <button
            onClick={toggleSidebar}
            className="hidden lg:block fixed bottom-4 left-4 z-50 p-2 bg-white border rounded-lg shadow-md hover:bg-gray-50 transition-colors"
            aria-label="Expand sidebar"
            data-testid="desktop-sidebar-expand"
          >
            <PanelLeftOpen className="w-5 h-5 text-gray-600" />
          </button>
        )}

        {/* Sidebar */}
        <Sidebar 
          isOpen={isMobileMenuOpen} 
          onClose={closeMobileMenu}
          isCollapsed={isSidebarCollapsed}
          onToggleCollapse={toggleSidebar}
        />

        {/* Main Content */}
        <div className="flex-1 flex flex-col overflow-hidden lg:ml-0">
          <main className="flex-1 overflow-auto">
            <Outlet />
          </main>
        </div>

        {typeof window !== 'undefined' && !window.navigator.userAgent.includes('Playwright') && <TanStackRouterDevtools />}
      </div>
    );
  },
});
