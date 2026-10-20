if (typeof HTMLCanvasElement !== 'undefined') {
  HTMLCanvasElement.prototype.getContext = ((contextType: string) => {
    if (contextType === '2d') {
      return {
        fillRect: () => {},
        clearRect: () => {},
        getImageData: (_x: number, _y: number, w: number, h: number) => ({
          data: new Uint8ClampedArray(w * h * 4),
        }),
        putImageData: () => {},
        createImageData: () => [],
        setTransform: () => {},
        drawImage: () => {},
        save: () => {},
        fillText: () => {},
        restore: () => {},
        beginPath: () => {},
        moveTo: () => {},
        lineTo: () => {},
        closePath: () => {},
        stroke: () => {},
        translate: () => {},
        scale: () => {},
        rotate: () => {},
        arc: () => {},
        fill: () => {},
        measureText: (text: string) => ({ width: text.length * 8, height: 14 }),
        transform: () => {},
        rect: () => {},
        clip: () => {},
      } as unknown as CanvasRenderingContext2D
    }
    return null
  }) as unknown as typeof HTMLCanvasElement.prototype.getContext
}
