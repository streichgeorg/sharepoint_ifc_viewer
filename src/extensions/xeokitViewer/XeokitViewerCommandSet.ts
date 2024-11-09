import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { BaseDialog } from '@microsoft/sp-dialog';

import "./XeokitViewer.module.scss";

import {
  Viewer,
  WebIFCLoaderPlugin,
  SectionPlanesPlugin,
  TreeViewPlugin,
  ContextMenu,
  math
} from './xeokit-sdk.es5.js'
import * as WebIFC from "./web-ifc-api.js";

import styles from './XeokitViewer.module.scss';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IXeokitViewerCommandSetProperties {
}

const LOG_SOURCE: string = 'XeokitViewerCommandSet';

const IfcAPI = new WebIFC.IfcAPI();

class ViewerDialog extends BaseDialog {
  private sceneModel: object;
  private modelPath: string;

  private viewer: any;
  private canvasContextMenu: any;
  private objectContextMenu: any;

  constructor(modelPath: string) {
    super();
    this.modelPath = modelPath;
  }

  private getKeyMap(cameraControl: any, input: any): {keyMap: object, helpText: string} {
    const helpText = `
      <div>
        <h3>Mouse:</h3>
        <ul>
          <li>Left Button: Rotate camera</li>
          <li>Right Button: Pan camera, Open context menu</li>
        </ul>

        <h3>Movement:</h3>
        <ul>
          <li>Left Arrow: Pan camera left</li>
          <li>Right Arrow: Pan camera right</li>
          <li>Up Arrow: Move camera forwards</li>
          <li>Down Arrow: Move camera backwards</li>
        </ul>
        
        <h3>Preset Views:</h3>
        <ul>
          <li>1: Right view</li>
          <li>2: Back view</li>
          <li>3: Left view</li>
          <li>4: Front view</li>
          <li>5: Top view</li>
          <li>6: Bottom view</li>
        </ul>
      </div>`;

    const keyMap = {};
    keyMap[cameraControl.PAN_LEFT] = [input.KEY_LEFT_ARROW];
    keyMap[cameraControl.PAN_RIGHT] = [input.KEY_RIGHT_ARROW];
    keyMap[cameraControl.DOLLY_FORWARDS] = [input.KEY_UP_ARROW];
    keyMap[cameraControl.DOLLY_BACKWARDS] = [input.KEY_DOWN_ARROW];
    keyMap[cameraControl.AXIS_VIEW_RIGHT] = [input.KEY_NUM_1];
    keyMap[cameraControl.AXIS_VIEW_BACK] = [input.KEY_NUM_2];
    keyMap[cameraControl.AXIS_VIEW_LEFT] = [input.KEY_NUM_3];
    keyMap[cameraControl.AXIS_VIEW_FRONT] = [input.KEY_NUM_4];
    keyMap[cameraControl.AXIS_VIEW_TOP] = [input.KEY_NUM_5];
    keyMap[cameraControl.AXIS_VIEW_BOTTOM] = [input.KEY_NUM_6];

    return {keyMap, helpText};
  }

  private initializeViewer(): void {
    Log.info(LOG_SOURCE, 'Initializing Viewer');

    const canvas = this.domElement.querySelector("#viewerCanvas")!;

    const viewer = new Viewer({
      canvasElement: canvas,
      transparent: true,
      numCachedSectionPlanes: 4
    });
    this.viewer = viewer;

    const cameraControl = this.viewer.cameraControl;
    cameraControl.navMode = "orbit";
    cameraControl.followPointer = true;
    const {keyMap, helpText} = this.getKeyMap(cameraControl, viewer.scene.input);
    cameraControl.keyMap = keyMap;

    const helpTextContainer = this.domElement.querySelector("#helpContainer")!;
    helpTextContainer.innerHTML = helpText;

    const treeView = new (TreeViewPlugin as any)(viewer, {
        containerElement: document.querySelector("#treeViewContainer")!,
        autoExpandDepth: 3,
        hierarchy: "containment"
    });

    Log.info(LOG_SOURCE, `TreeView initialized: ${treeView}`);

    const sectionPlanes = new (SectionPlanesPlugin as any)(viewer, {
      overviewCanvasId: this.domElement.querySelector("#sectionPlaneCanvas")!.id,
      overviewVisible: true
    });

    Log.info(LOG_SOURCE, `SectionPlanes initialized: ${sectionPlanes}`);

    const canvasContextMenu = new ContextMenu({
      enabled: true,
      context: {
        viewer: viewer
      },
      items: [
        [
          {
            title: "Hide All",
            getEnabled: function (context) {
              return (context.viewer.scene.numVisibleObjects > 0);
            },
            doAction: function (context) {
              context.viewer.scene.setObjectsVisible(context.viewer.scene.visibleObjectIds, false);
            }
          },
          {
            title: "Show All",
            getEnabled: function (context) {
              const scene = context.viewer.scene;
              return (scene.numVisibleObjects < scene.numObjects);
            },
            doAction: function (context) {
              const scene = context.viewer.scene;
              scene.setObjectsVisible(scene.objectIds, true);
              scene.setObjectsXRayed(scene.xrayedObjectIds, false);
              scene.setObjectsSelected(scene.selectedObjectIds, false);
            }
          }
        ],
        [
          {
            title: "View Fit All",
            doAction: function (context) {
              context.viewer.cameraFlight.flyTo({
                aabb: context.viewer.scene.getAABB()
              });
            }
          }
        ]
      ]
    });
    this.canvasContextMenu = canvasContextMenu;

    const objectContextMenu = new ContextMenu({
      items: [
        [
          {
            title: "View Fit",
            doAction: function (context) {
              const viewer = context.viewer;
              const scene = viewer.scene;
              const entity = context.entity;
              viewer.cameraFlight.flyTo({
                aabb: entity.aabb,
                duration: 0.5
              }, () => {
                setTimeout(function () {
                  scene.setObjectsHighlighted(scene.highlightedObjectIds, false);
                }, 500);
              });
            }
          },
          {
            title: "View Fit All",
            doAction: function (context) {
              const scene = context.viewer.scene;
              context.viewer.cameraFlight.flyTo({
                projection: "perspective",
                aabb: scene.getAABB(),
                duration: 0.5
              });
            }
          },
          {
            title: "Show in Tree",
            doAction: function (context) {
              const objectId = context.entity.id;
              context.treeViewPlugin.showNode(objectId);
            }
          }
        ],
        [
          {
            title: "Hide",
            getEnabled: function (context) {
              return context.entity.visible;
            },
            doAction: function (context) {
              context.entity.visible = false;
            }
          },
          {
            title: "Hide Others",
            doAction: function (context) {
              const viewer = context.viewer;
              const scene = viewer.scene;
              const entity = context.entity;
              const metaObject = viewer.metaScene.metaObjects[entity.id];
              if (!metaObject) {
                return;
              }
              scene.setObjectsVisible(scene.visibleObjectIds, false);
              scene.setObjectsXRayed(scene.xrayedObjectIds, false);
              scene.setObjectsSelected(scene.selectedObjectIds, false);
              scene.setObjectsHighlighted(scene.highlightedObjectIds, false);
              metaObject.withMetaObjectsInSubtree((metaObject) => {
                const entity = scene.objects[metaObject.id];
                if (entity) {
                  entity.visible = true;
                }
              });
            }
          },
          {
            title: "Hide All",
            getEnabled: function (context) {
              return (context.viewer.scene.numVisibleObjects > 0);
            },
            doAction: function (context) {
              context.viewer.scene.setObjectsVisible(context.viewer.scene.visibleObjectIds, false);
            }
          },
          {
            title: "Show All",
            getEnabled: function (context) {
              const scene = context.viewer.scene;
              return (scene.numVisibleObjects < scene.numObjects);
            },
            doAction: function (context) {
              const scene = context.viewer.scene;
              scene.setObjectsVisible(scene.objectIds, true);
            }
          }
        ],
        [
          {
            title: "Create Section Plane",
            doAction: function (context) {
              const pickResult = context.viewer.scene.pick({
                canvasPos: context.canvasPos,
                pickSurface: true  // <<------ This causes picking to find the intersection point on the entity
              });

              if (pickResult && pickResult.worldNormal) { // Disallow SectionPlanes on point clouds, because points don't have normals

                if (pickResult.entity) {
                  if (!pickResult.entity.isObject) {
                    return;
                  }
                }

                const sectionPlane = sectionPlanes.createSectionPlane({
                  pos: pickResult.worldPos,
                  dir: math.mulVec3Scalar(pickResult.worldNormal, -1)
                });

                sectionPlanes.showControl(sectionPlane.id);
              }
            }
          },
          {
            title: "Clear Section Planes",
            doAction: function () {
              sectionPlanes.clear();
            }
          }
        ]
      ],
      enabled: true
    });

    const getCanvasPosFromEvent = function (event) {
      const canvasPos: number[] = [];
      if (!event) {
        event = window.event;
        canvasPos[0] = event.x;
        canvasPos[1] = event.y;
      } else {
        let element = event.target;
        let totalOffsetLeft = 0;
        let totalOffsetTop = 0;
        let totalScrollX = 0;
        let totalScrollY = 0;
        while (element.offsetParent) {
          totalOffsetLeft += element.offsetLeft;
          totalOffsetTop += element.offsetTop;
          totalScrollX += element.scrollLeft;
          totalScrollY += element.scrollTop;
          element = element.offsetParent;
        }
        canvasPos[0] = event.pageX + totalScrollX - totalOffsetLeft;
        canvasPos[1] = event.pageY + totalScrollY - totalOffsetTop;
      }
      return canvasPos;
    };

    viewer.scene.canvas.canvas.addEventListener('contextmenu', function (event) {
      const canvasPos = getCanvasPosFromEvent(event);
      const hit = viewer.scene.pick({
        canvasPos
      });
      if (hit && hit.entity.isObject) {
        objectContextMenu.context = { // Must set context before showing menu
          viewer,
          treeViewPlugin: treeView,
          entity: hit.entity,
          canvasPos,
        };
        (objectContextMenu as any).show(event.pageX, event.pageY);
      } else {
        canvasContextMenu.context = { // Must set context before showing menu
          viewer
        };
        (canvasContextMenu as any).show(event.pageX, event.pageY);
      }
      event.preventDefault();
    });
    this.objectContextMenu = objectContextMenu;

    Log.info(LOG_SOURCE, 'Initializing WebIFCLoaderPlugin');

    const ifcLoader = new (WebIFCLoaderPlugin as any)(viewer, {
      WebIFC,
      IfcAPI
    });

    Log.info(LOG_SOURCE, 'Loading model...');

    this.sceneModel = ifcLoader.load({
      id: "model",
      src: this.modelPath,
      edges: true,
      backfaces: true,
      loadMetadata: true
    });

    Log.info(LOG_SOURCE, `Model loaded: ${this.sceneModel}`);
  }

  public render(): void {
    this.domElement.innerHTML = `<div class="${styles.viewerContainer}">
      <div class="${styles.helpButton}">
        <span>Help</span>
      </div>
      <div id="helpContainer" class="${styles.helpContainer}"></div>
      <div id="treeViewContainer" class="${styles.treeViewContainer}"></div>
      <canvas id="viewerCanvas" class="${styles.viewerCanvas}" tabindex="0"></canvas>
      <canvas id="sectionPlaneCanvas" class="${styles.sectionPlaneCanvas}"></canvas>
    </div>`;

    Log.info(LOG_SOURCE, "Waiting for DOM to be ready...");

    // Wait until the element is actually added to the DOM because otherwise the split
    // plane plugin won't find the canvas element.
    const observer = new MutationObserver((mutations) => {
      if (document.querySelector("#sectionPlaneCanvas")) {
        this.initializeViewer();
        observer.disconnect();
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  public onAfterClose(): void {
    if (this.viewer) {
      this.viewer.destroy();
    }
    if (this.canvasContextMenu) {
      this.canvasContextMenu.destroy();
    }
    if (this.objectContextMenu) {
      this.objectContextMenu.destroy();
    }
  }
}

export default class XeokitViewerCommandSet extends BaseListViewCommandSet<IXeokitViewerCommandSetProperties> {
  private ifcAPIInitialized: Promise<void>;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized XeokitViewerCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    IfcAPI.SetWasmPath("https://cdn.jsdelivr.net/npm/web-ifc@0.0.51/");
    this.ifcAPIInitialized = (IfcAPI as any).Init();

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        if (this.context.listView.selectedRows && this.context.listView.selectedRows.length === 1) {
          const fileRef = this.context.listView.selectedRows[0].getValueByName("FileRef");
          const viewerDialog = new ViewerDialog(fileRef);

          this.ifcAPIInitialized.then(() => {
            return viewerDialog.show();
          }).catch(console.error);
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');

    if (compareOneCommand && this.context.listView.selectedRows?.length === 1) {
      const item = this.context.listView.selectedRows[0]
      const fileType = item.getValueByName(".fileType");
      if (fileType && fileType === "ifc") {
        compareOneCommand.visible = true;
      }
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
