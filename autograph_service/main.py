"""GUI-first entrypoint for AutoGraph."""

from autograph_service.gui import launch_gui


def main():
    """Launch the Tkinter GUI."""
    launch_gui()


if __name__ == "__main__":
    main()
