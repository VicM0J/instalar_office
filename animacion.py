import tkinter as tk
import math

class InfinityDownloadAnimation:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargando Office...")
        self.canvas = tk.Canvas(root, width=600, height=400, bg="black")
        self.canvas.pack()

        self.center_x = 300
        self.center_y = 200
        self.radius = 80
        self.phase = 0
        self.speed = 0.05
        self.trail_length = 50

        self.trail_points = []
        self.arrow = self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="orange")
        self.trail = self.canvas.create_line([], fill="orange", width=3, smooth=True)

        self.box = self.canvas.create_rectangle(270, 330, 330, 350, outline="white", fill="gray")
        self.descending = False
        self.descend_step = 0
        self.max_descend_steps = 60

        self.animate()

    def infinity_path(self, t):
        a = self.radius
        x = a * math.sin(t)
        y = a * math.sin(t) * math.cos(t)
        return self.center_x + x * 1.5, self.center_y + y

    def animate(self):
        if not self.descending:
            x, y = self.infinity_path(self.phase)
            self.trail_points.append((x, y))
            if len(self.trail_points) > self.trail_length:
                self.trail_points.pop(0)

            flat_points = [coord for point in self.trail_points for coord in point]
            if len(flat_points) >= 4:
                self.canvas.coords(self.trail, *flat_points)

            angle = self.phase
            dx = 10 * math.cos(angle)
            dy = 10 * math.sin(angle)
            self.canvas.coords(
                self.arrow,
                x, y,
                x - dy, y + dx,
                x + dy, y - dx
            )

            self.phase += self.speed
            if self.phase >= 4 * math.pi:
                self.descending = True
                self.descend_step = 0
        else:
            if self.trail_points:
                self.trail_points.pop(0)
                flat_points = [coord for point in self.trail_points for coord in point]
                self.canvas.coords(self.trail, *flat_points)

            if self.descend_step <= self.max_descend_steps:
                progress = self.descend_step / self.max_descend_steps
                x = self.center_x
                y = self.center_y + progress * 130
                self.canvas.coords(
                    self.arrow,
                    x, y,
                    x - 10, y + 15,
                    x + 10, y + 15
                )
                self.descend_step += 1
            else:
                self.phase = 0
                self.trail_points.clear()
                self.descending = False

        self.root.after(30, self.animate)

# Ejecutar solo en un entorno con interfaz gráfica (no en servidores sin display)
if __name__ == "__main__":
    root = tk.Tk()
    app = InfinityDownloadAnimation(root)
    root.mainloop()
