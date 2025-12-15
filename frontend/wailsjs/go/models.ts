export namespace main {
	
	export class BackupAvailability {
	    available: boolean;
	    reason?: string;
	    message: string;
	    path?: string;
	    notebookCount?: number;
	    size?: string;
	
	    static createFrom(source: any = {}) {
	        return new BackupAvailability(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.available = source["available"];
	        this.reason = source["reason"];
	        this.message = source["message"];
	        this.path = source["path"];
	        this.notebookCount = source["notebookCount"];
	        this.size = source["size"];
	    }
	}
	export class ExportResult {
	    success: boolean;
	    message: string;
	    exportedPath?: string;
	
	    static createFrom(source: any = {}) {
	        return new ExportResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.success = source["success"];
	        this.message = source["message"];
	        this.exportedPath = source["exportedPath"];
	    }
	}
	export class FileItem {
	    name: string;
	    path: string;
	    isDir: boolean;
	    children?: FileItem[];
	    level: number;
	
	    static createFrom(source: any = {}) {
	        return new FileItem(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.name = source["name"];
	        this.path = source["path"];
	        this.isDir = source["isDir"];
	        this.children = this.convertValues(source["children"], FileItem);
	        this.level = source["level"];
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}
	export class NotebookInfo {
	    id: string;
	    name: string;
	    path: string;
	    lastModified: string;
	    isCurrentlyViewed: boolean;
	
	    static createFrom(source: any = {}) {
	        return new NotebookInfo(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.id = source["id"];
	        this.name = source["name"];
	        this.path = source["path"];
	        this.lastModified = source["lastModified"];
	        this.isCurrentlyViewed = source["isCurrentlyViewed"];
	    }
	}
	export class VersionInfo {
	    version: string;
	    oneNoteInstalled: boolean;
	    oneNoteVersion: string;
	
	    static createFrom(source: any = {}) {
	        return new VersionInfo(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.version = source["version"];
	        this.oneNoteInstalled = source["oneNoteInstalled"];
	        this.oneNoteVersion = source["oneNoteVersion"];
	    }
	}

}

