interface BannerOptions {
  message: string
}

interface PluginRenderContext {
  compileNode: (node: unknown) => Promise<unknown>
}

const lifecyclePlugin = {
  name: 'lifecycle-banner',
  render(options: BannerOptions, context: PluginRenderContext) {
    return context.compileNode({
      children: [
        {
          style: {
            color: '#1D4ED8',
            fontWeight: 'bold',
          },
          text: `Plugin lifecycle: ${options.message}`,
          type: 'text',
        },
      ],
      style: {
        backgroundColor: '#EFF6FF',
        borderLeft: {
          color: '#2563EB',
          style: 'single',
          width: '2pt',
        },
      },
      type: 'paragraph',
    })
  },
}

export default lifecyclePlugin
