        # ERP-Warehouse-Assistant — Agent Guide

        This project lives inside **Derle**, an on-device mobile IDE. Your terminal, build tools,
        and this project all run inside a **proot Ubuntu Linux** environment on an Android device
        (no root required). Standard Linux tooling is available — `bash`, `git`, `apt`, plus any
        toolchains the user installed from Derle's **Components / SDKs** screen.

        Detected project type: **Python**.

        ## Running this project
        The **Run** button executes `.derle/run.sh` (its current command is `python3 app.py`). To change how this project runs, **edit `.derle/run.sh`** — it is the single source of truth, and the app reads it every time Run is pressed. You can put any shell logic there.

        - To change how the project runs, **edit `.derle/run.sh`** — it is the single source of
  truth for the run command, and the app reads it every time Run is pressed.
- You can put any shell logic in it (build steps, compiler flags, env vars).
- If this project's own docs (`README.md` and the other markdown at the root) describe a
  different way to start it, make `.derle/run.sh` match them — otherwise the Run button
  and the documented command disagree.
- The user can restore the default from the app's **Run Settings** tab.

        ## Environment notes
        - Installed toolchains live under `/home/root/` and expose their binaries on `PATH` via
          `/usr/local/bin`. Once installed, compilers (`gcc`, `g++`, `clang`), `node`/`npm`, and
          Python (as `python3.x`) are directly runnable.
        - A few SDKs (JDK, Flutter, Android SDK) are wired into their language's build flow rather
          than the global `PATH`. Prefer the project's build tool (the `gradlew`, `flutter`, or
          `mvnw` wrapper) so the right toolchain and env vars are used.
        - C/C++ builds output to `build/`. Keep generated artifacts there.
        - Install more dependencies with the environment's package managers (`apt`, `pip`, `npm`)
          or from Derle's Components screen.
        - Long-running commands (dev servers, builds) survive backgrounding — Derle keeps the
          shell alive with a foreground service.

        ## Conventions
        - Don't change the build-system flavor (CMake ↔ Makefile, Gradle, Vite, …) unless asked.
        - Keep edits minimal and match the surrounding code style.
