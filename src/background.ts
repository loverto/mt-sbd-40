import {
  app,
  globalShortcut,
  protocol
} from 'electron'
import {
  createProtocol,
  installVueDevtools
} from 'vue-cli-plugin-electron-builder/lib'
import db from '#/datastore'

import beforeOpen from '~/main/utils/beforeOpen'
import updateChecker from '~/main/utils/updateChecker'
import fixPath from 'fix-path'
import { getUploadFiles } from '~/main/utils/handleArgv'
import bus from '~/main/utils/eventBus'


const isDevelopment = process.env.NODE_ENV !== 'production'
protocol.registerSchemesAsPrivileged([{ scheme: 'picgo', privileges: { secure: true, standard: true } }])

beforeOpen()

// fix the $PATH in macOS
fixPath()


const gotTheLock = app.requestSingleInstanceLock()

if (!gotTheLock) {
  app.quit()
} else {
  app.on('second-instance', (event, commandLine, workingDirectory) => {
  })
}

if (process.platform === 'win32') {
  app.setAppUserModelId('com.molunerfinn.picgo')
}

if (process.env.XDG_CURRENT_DESKTOP && process.env.XDG_CURRENT_DESKTOP.includes('Unity')) {
  process.env.XDG_CURRENT_DESKTOP = 'Unity'
}

app.on('ready', async () => {
  createProtocol('picgo')
  if (isDevelopment && !process.env.IS_TEST) {
    // Install Vue Devtools
    try {
      await installVueDevtools()
    } catch (e) {
      console.error('Vue Devtools failed to install:', e.toString())
    }
  }
  if (process.platform === 'darwin' || process.platform === 'win32') {
  }
  db.set('needReload', false)
  updateChecker()
  // 不需要阻塞
  process.nextTick(() => {
  })
  if (process.env.NODE_ENV !== 'development') {

  }
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {

})

app.on('will-quit', () => {
  globalShortcut.unregisterAll()
  bus.removeAllListeners()
})

app.setLoginItemSettings({
  openAtLogin: db.get('settings.autoStart') || false
})

// Exit cleanly on request from parent process in development mode.
if (isDevelopment) {
  if (process.platform === 'win32') {
    process.on('message', data => {
      if (data === 'graceful-exit') {
        app.quit()
      }
    })
  } else {
    process.on('SIGTERM', () => {
      app.quit()
    })
  }
}

/**
 * Auto Updater
 *
 * Uncomment the following code below and install `electron-updater` to
 * support auto updating. Code Signing with a valid certificate is required.
 * https://simulatedgreg.gitbooks.io/electron-vue/content/en/using-electron-builder.html#auto-updating
*/

// import { autoUpdater } from 'electron-updater'

// autoUpdater.on('update-downloaded', () => {
//   autoUpdater.quitAndInstall()
// })

// app.on('ready', () => {
//   if (process.env.NODE_ENV === 'production') {
//     autoUpdater.checkForUpdates()
//   }
// })
