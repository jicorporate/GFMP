from PIL import Image
import numpy as np

def trace_contours():
    img = Image.open('scratch/extracted_foreground.png').convert('RGBA')
    arr = np.array(img)
    mask = (arr[:, :, 3] > 140).astype(np.uint8)
    h, w = mask.shape

    # Marching squares or 8-connected contour following
    # Moore-Neighbor Tracing
    # Directions: (dx, dy)
    # 0: E, 1: SE, 2: S, 3: SW, 4: W, 5: NW, 6: N, 7: NE
    dirs = [(1, 0), (1, 1), (0, 1), (-1, 1), (-1, 0), (-1, -1), (0, -1), (1, -1)]

    # We want to find all outer and hole contours
    visited_edges = set()
    contours = []

    # Pad image by 1
    padded = np.zeros((h + 2, w + 2), dtype=np.uint8)
    padded[1:-1, 1:-1] = mask

    for y in range(1, h + 1):
        for x in range(1, w + 1):
            if padded[y, x] == 1:
                # Check 4 borders to start tracing
                # If top neighbor is 0 and edge not visited
                if padded[y - 1, x] == 0 and (x, y, 6) not in visited_edges:
                    # Start tracing outer contour
                    curr_x, curr_y = x, y
                    # We arrived from (x, y-1), so entry dir was 2 (from North)
                    # Backtrack direction: looking from curr towards entry
                    # In Moore tracing, check counterclockwise starting from backtrack
                    contour = []
                    start_x, start_y = curr_x, curr_y
                    start_dir = 6

                    # Trace loop
                    curr_dir = 6
                    max_steps = 10000
                    steps = 0
                    while steps < max_steps:
                        steps += 1
                        contour.append((curr_x - 1, curr_y - 1)) # back to 0-based
                        visited_edges.add((curr_x, curr_y, curr_dir))

                        # Search around curr_x, curr_y starting from (curr_dir + 5) % 8
                        search_start = (curr_dir + 5) % 8
                        found = False
                        for i in range(8):
                            d = (search_start + i) % 8
                            nx = curr_x + dirs[d][0]
                            ny = curr_y + dirs[d][1]
                            if padded[ny, nx] == 1:
                                curr_x = nx
                                curr_y = ny
                                curr_dir = d
                                found = True
                                break

                        if not found:
                            break
                        if curr_x == start_x and curr_y == start_y and curr_dir == start_dir:
                            break

                    if len(contour) > 5:
                        contours.append(contour)

    print(f"Found {len(contours)} contours")
    return contours

if __name__ == '__main__':
    contours = trace_contours()

