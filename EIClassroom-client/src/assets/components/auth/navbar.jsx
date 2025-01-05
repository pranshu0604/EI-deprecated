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
