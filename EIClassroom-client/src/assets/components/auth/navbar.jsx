import { useNavigate } from "react-router-dom";
import { useState, useEffect } from "react";
export default function Navbar() {
  const navigate = useNavigate();
  const [theme, setTheme] = useState(
    typeof window !== "undefined" && window.matchMedia("(prefers-color-scheme: dark)").matches
      ? "dark"
      : "light"
  );
  const handleThemeToggle = () => {
    const newTheme = theme === "dark" ? "light" : "dark";
    setTheme(newTheme);
    document.documentElement.classList.toggle("dark", newTheme === "dark");
    localStorage.setItem("theme", newTheme);
  };
  useEffect(() => {
    const savedTheme = localStorage.getItem("theme");
    if (savedTheme) {
      setTheme(savedTheme);
      document.documentElement.classList.toggle("dark", savedTheme === "dark");
    }
  }, []);

  return (
    <div className="chakra-petch-medium dark:text-white bg-transparent backdrop-blur-lg m-4 w-[92%] md:w-[98%] fixed z-30 rounded-lg border-stone-700 border-2">
      <nav className="flex justify-between items-center px-3 py-3">
        <div className="flex-1">
          <h1 className="md:text-2xl text-nowrap text-xl px-3 font-bold text-black dark:text-white">
            <a href="/dashboard">EI Classroom Portal</a>
          </h1>
        </div>

        <div className="flex-1 hidden md:flex justify-center">
        <div className="flex items-center justify-center">
          <label className="relative inline-flex items-center cursor-pointer mr-2">
            <input
              type="checkbox"
              className="sr-only"
              checked={theme === "dark"}
              onChange={handleThemeToggle}
            />
            <div className="w-12 h-6 bg-gray-200 dark:bg-gray-800 rounded-full shadow-inner"></div>
            <div
              className={`absolute w-6 h-6 bg-white rounded-full shadow transition-transform duration-300 ease-in-out ${
                theme === "dark" ? "translate-x-6" : "translate-x-0"
              }`}
            ></div>
          </label>
          <span className="text-xl">{theme === "dark" ? "🌙" : "☀️"}</span>
        </div>
        </div>

        <div className="flex-1 flex justify-end">
          <button
            className="mx-2 py-1 px-4 bg-blue-500 rounded-sm"
            onClick={() => {
                localStorage.clear();
                navigate("/")
            }}
          >
            Sign Out
          </button>
        </div>
      </nav>
    </div>
  );
}
