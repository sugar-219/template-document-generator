"""Short, cancellable Tk transitions; all frames run on the UI thread."""
import time


class Motion:
    def __init__(self, widget, enabled=True):
        self.widget = widget
        self.enabled = enabled
        self.timer = None
        self.destination = None
        self.draw = None

    def cancel(self):
        if self.timer is not None:
            self.widget.after_cancel(self.timer)
            self.timer = None

    def start(self, origin, destination, draw, duration=180):
        self.cancel()
        self.destination, self.draw = destination, draw
        if not self.enabled or origin == destination:
            draw(destination)
            return
        started = time.monotonic()
        def frame():
            self.timer = None
            fraction = min(1, (time.monotonic() - started) * 1000 / duration)
            eased = 1 - (1 - fraction) ** 3
            draw(tuple(first + (last-first) * eased for first, last in zip(origin, destination)))
            if fraction < 1:
                self.timer = self.widget.after(16, frame)
        frame()

    def finish(self):
        self.cancel()
        if self.draw is not None:
            self.draw(self.destination)
