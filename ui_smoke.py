"""Проверка интерфейса в CI (Windows): сколько строятся и рисуются экраны, не тормозят ли анимации."""
import sys
import time

import bohemia_gui as g
import motion


def main():
    t0 = time.perf_counter()
    app = g.App()
    app.update()
    print(f'start: {time.perf_counter() - t0:.2f} s')
    t = time.perf_counter()
    while time.perf_counter() - t < 1.5:          # даём построиться экранам в фоне
        app.update()
        time.sleep(0.01)
    res = {}
    for key in ('docs', 'photos'):
        t = time.perf_counter()
        app.show(key)
        app.update()
        res[key] = time.perf_counter() - t
        print(f'show {key}: {res[key] * 1000:.0f} ms')
    d = app.built['docs']
    app.show('docs')
    app.update()
    for name, fn in (('hover', lambda: [b.hover.to(1) for b in motion_buttons(app)]),
                     ('drop', lambda: d.docs.field.dropped(40_000_000))):
        fn()
        t = time.perf_counter()
        frames = 0
        while motion._active and time.perf_counter() - t < 3:
            app.update()
            frames += 1
        dur = time.perf_counter() - t
        print(f'{name}: {dur:.2f} s, loop passes {frames}, push={motion.PUSH}')
    app.destroy()
    if max(res.values()) > 1.0:
        sys.exit('screen switch is too slow')


def motion_buttons(app):
    out = []

    def walk(w):
        for c in w.winfo_children():
            if isinstance(c, motion.MotionButton):
                out.append(c)
            walk(c)
    walk(app)
    return out


if __name__ == '__main__':
    main()
