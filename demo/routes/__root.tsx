import { createRootRoute, Outlet } from "@tanstack/react-router";
import { TanStackRouterDevtools } from "@tanstack/react-router-devtools";
import { useEffect, useState } from "react";
import { Sidebar } from "../components/Sidebar";
import { MobileMenuToggle } from "../components/MobileMenuToggle";
import "../index.css";

export const Route = createRootRoute({
  component: () => {
    const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

    const toggleMobileMenu = () => {
      setIsMobileMenuOpen(!isMobileMenuOpen);
    };

    const closeMobileMenu = () => {
      setIsMobileMenuOpen(false);
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

        {/* Sidebar */}
        <Sidebar isOpen={isMobileMenuOpen} onClose={closeMobileMenu} />

        {/* Main Content */}
        <div className="flex-1 flex flex-col overflow-hidden lg:ml-0">
          <main className="flex-1 overflow-auto p-6 pt-16 lg:pt-6">
            <Outlet />
          </main>
        </div>

        {typeof window !== 'undefined' && !window.navigator.userAgent.includes('Playwright') && <TanStackRouterDevtools />}
      </div>
    );
  },
});
