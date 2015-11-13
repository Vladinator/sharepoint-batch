/*!
 * SharePointBatch
 * Copyright 2015 Alex Pedersen
 * Licensed under the MIT license
 * https://github.com/Vladinator89/sharepoint-batch
 */
(function(){
	'use strict';

	/**
	 * batch constructor
	 */
	function batch(options) {
		extend(this, {
			options: extend({
				url: '',
				digest: ''
			}, options),
			guid: createGUID(),
			jobs: [],
			/**
			 * public interface
			 */
			append: append,
			remove: remove,
			payload: payload,
			toString: payload,
			send: send,
			spawn: spawn
		});

		/**
		 * append a job to the batch
		 *
		 * @return integer index of the job
		 */
		function append(options) {
			options = extend({
				guid: createGUID(),
				method: 'GET',
				url: '',
				headers: {},
				changesets: [],
				args: {}
			}, options);

			/**
			 * public interface
			 */
			extend(options, {
				payload: payload,
				toString: payload
			});

			// payload
			var data = [];

			// must contain at least one element (null means no payload - for example used by DELETE requests)
			if (!isArray(options.changesets) || !options.changesets.length) {
				options.changesets = [null];
			}

			// call appropriate handler
			(options.method !== 'GET' ? post : get).call(this);

			// store job in our array
			return this.jobs.push(options) - 1;

			/**
			 * handle GET related requests
			 */
			function get() {
				data.push('--batch_' + this.guid);

				data.push('Content-Type: application/http');
				data.push('Content-Transfer-Encoding: binary');
				data.push('');

				// TODO: remove this loop and just perform the contents once? can you actually have multiple "gets" like this? needs testing before removing/altering!
				for (var i = 0; i < options.changesets.length; i++) {
					data.push(options.method + ' ' + options.url + toParams(options.params) + ' HTTP/1.1');
					data.push('Accept: application/json;odata=verbose');

					for (var key in options.headers) {
						if (options.headers.hasOwnProperty(key)) {
							data.push(key + ': ' + options.headers[key]);
						}
					}

					data.push('');
				}
			}

			/**
			 * handle POST related requests (anything not GET defaults to this method)
			 */
			function post() {
				var boundary = 'changeset_' + createGUID();

				data.push('--batch_' + this.guid);

				data.push('Content-Type: multipart/mixed; boundary="' + boundary + '"');
				data.push('Content-Transfer-Encoding: binary');
				data.push('');

				for (var i = 0; i < options.changesets.length; i++) {
					var changeset = options.changesets[i];

					data.push('--' + boundary);
					data.push('Content-Type: application/http');
					data.push('Content-Transfer-Encoding: binary');
					data.push('');

					data.push(options.method + ' ' + options.url + toParams(options.params) + ' HTTP/1.1');
					data.push('Accept: application/json;odata=verbose');
					data.push('Content-Type: application/json;odata=verbose');

					var headers = extend({}, options.headers);
					extend(headers, changeset.headers);

					for (var key in headers) {
						if (headers.hasOwnProperty(key)) {
							data.push(key + ': ' + headers[key]);
						}
					}

					data.push('');

					if (changeset.data !== null) {
						data.push(JSON.stringify(changeset.data));
						data.push('');
					}
				}

				data.push('--' + boundary + '--');
			}

			/**
			 * get the payload for this job
			 *
			 * @return string
			 */
			function payload() {
				return data.join('\r\n');
			}
		}

		/**
		 * remove a job by index
		 *
		 * @return object removes and returns the job
		 */
		function remove(index) {
			if (typeof index === 'number') {
				return this.jobs.splice(index, 1);
			}
		}

		/**
		 * build batch payload
		 *
		 * @return string
		 */
		function payload() {
			var s = '';

			for (var i = 0; i < this.jobs.length; i++) {
				var job = this.jobs[i];

				if (job) {
					s += job.payload() + '\r\n';
				}
			}

			return s;
		}

		/**
		 * send request to the $batch handler
		 */
		function send(options) {
			options = extend({
				method: 'POST',
				url: this.options.url + '/_api/$batch',
				headers: {
					'X-RequestDigest': this.options.digest,
					'Content-Type': 'multipart/mixed; boundary="batch_' + this.guid + '"'
				},
				data: this.payload() + '--batch_' + this.guid + '--'
			}, options);

			// backup methods before overriding
			var backup = extend({}, options);

			// properly parse the loaded data
			options.done = function(xhr) {
				xhr.responseJSON = parseResponse.call(this, xhr.responseText);
				decorateJobs.call(this, xhr.responseJSON);

				Array.prototype.push.call(arguments, xhr.responseJSON);

				if (typeof backup.done === 'function') {
					backup.done.apply(this, arguments);
				}
			};

			// properly parse the error message
			options.fail = function(xhr) {
				xhr.responseJSON = parseResponse.call(this, xhr.responseText);
				decorateJobs.call(this, xhr.responseJSON);

				Array.prototype.push.call(arguments, xhr.responseJSON);

				if (typeof backup.fail === 'function') {
					backup.fail.apply(this, arguments);
				}
			};

			// properly parse the data/error
			options.always = function(xhr) {
				Array.prototype.push.call(arguments, xhr.responseJSON);

				if (typeof backup.always === 'function') {
					backup.always.apply(this, arguments);
				}
			};

			// return handle for the ajax request
			return ajax.call(this, options);

			/**
			 * safely read the response and conver it to a json object
			 * the input can either be a simple json string, or something complex like a
			 * boundary separated response with batchresults and changeresults
			 *
			 * @return object
			 */
			function parseResponse(raw) {
				if (typeof raw !== 'string') {
					return null;
				}

				// try parsing this, if it fails, it has to be multipart/mixed payload
				try {
					return JSON.parse(raw);
				} catch (e) {
				}

				// enum where we are in the payload
				var LEVEL = {
					UNKNOWN: 0,
					HEADERS: 1,
					REQUEST: 2,
					REQUEST_HEADERS: 3,
					REQUEST_BODY: 4,
					EOF: 5
				};

				// split the multipart/mixed into lines
				var lines = raw.split(/\r\n/),
					results = [],
					temp = undefined,
					cwo = undefined,
					level = LEVEL.UNKNOWN;

				// parse each line
				for (var i = 0; i < lines.length; i++) {
					var line = lines[i];

					if (/^--batchresponse_.+--$/i.test(line)) {
						if (temp !== undefined) {
							temp.data = parseResponse.call(this, temp.data);
							results.push(temp);
							temp = undefined;
						}
						level = LEVEL.EOF;
						break;

					} else if (/^--batchresponse_.+$/i.test(line)) {
						if (temp !== undefined) {
							temp.data = parseResponse.call(this, temp.data);
							results.push(temp);
						}

						temp = {
							headers: {
							},
							http: {
								status: 0,
								statusText: ''
							},
							data: undefined
						};
						cwo = temp;
						level = LEVEL.HEADERS;

					} else if (level === LEVEL.REQUEST_BODY) {
						if (cwo.data === undefined) {
							cwo.data = '';
						}
						cwo.data += line;

					} else if (/^HTTP\/1\.1\s+(\d+)\s+(.+)$/i.test(line)) {
						if (level === LEVEL.REQUEST) {
							var http = line.match(/^HTTP\/1\.1\s+(\d+)\s+(.+)$/i);
							cwo.http.status = parseInt(http[1], 10);
							cwo.http.statusText = http[2];
							level = LEVEL.REQUEST_HEADERS;
						}

					} else if (/^.+:\s*.+$/i.test(line)) {
						if (level === LEVEL.HEADERS || level === LEVEL.REQUEST_HEADERS) {
							var parts = line.split(/:/),
								key = parts.shift(1).trim(),
								value = parts.join(':').trim();

							cwo.headers[key] = value;
						}

					} else if (/^[\s\r\n]*$/i.test(line)) {
						if (level === LEVEL.HEADERS) {
							level = LEVEL.REQUEST;
						} else if (level === LEVEL.REQUEST) {
							level = LEVEL.REQUEST_HEADERS;
						} else if (level === LEVEL.REQUEST_HEADERS) {
							level = LEVEL.REQUEST_BODY;
						}
					}
				}

				// return the results
				return results;
			}

			/**
			 * assigns the job objects their results
			 *
			 * @return object
			 */
			function decorateJobs(results) {
				var index = 0;

				for (var i = 0; i < this.jobs.length; i++) {
					var job = this.jobs[i];

					if (!isArray(job.changesetResults)) {
						job.changesetResults = [];
					} else {
						while (job.changesetResults.length) {
							job.changesetResults.pop();
						}
					}

					for (var j = 0; j < job.changesets.length; j++) {
						var result = results[index++];

						if (isObject(result, true)) {
							result.job = job;
							result.changeset = j;
						}

						job.changesetResults.push(result);
					}
				}
			}
		}

		/**
		 * spawn a new instance using the same options as the current object
		 *
		 * @return object
		 */
		function spawn(options) {
			return new batch(extend(extend({}, this.options), options));
		}
	}

	/**
	 * create random guid
	 *
	 * @return string
	 */
	function createGUID() {
		return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
			var r = Math.random() * 16 | 0;
			return (c == 'x' ? r : (r & 0x3 | 0x8)).toString(16);
		});
	}

	/**
	 * convert object to a url query string
	 *
	 * @return string
	 */
	function toParams(o) {
		var s = '';

		if (isObject(o)) {
			for (var key in o) {
				var v = '';
				try {
					v = encodeURIComponent(o[key]);
				} catch (e) {
				}
				s += (s === '' ? '?' : '&') + key + '=' + v;
			}
		}

		return s;
	}

	/**
	 * extend(true, {a}, {b})
	 * where object b is copied into object a children objects will
	 * also be copied (creating a clone)
	 *
	 * @return object a
	 */
	function extend(arg1, arg2, arg3) {
		var deep = arg1 === true,
				dst = deep ? arg2 : arg1,
				src = deep ? arg3 : arg2;

		if (!isObject(dst) && typeof dst !== 'function') {
			dst = {};
		}

		if (isObject(src)) {
			var dstVal, srcVal, srcIsArray;

			for (var key in src) {
				srcVal = src[key];
				dstVal = dst[key];
				srcIsArray = isArray(srcVal);

				if (srcVal === dstVal) {
					continue;
				}

				if (deep && srcVal && (isObject(srcVal) || srcIsArray)) {
					extend(deep, dst[key], srcVal);
				} else if (srcVal !== undefined) {
					dst[key] = srcVal;
				}
			}
		}

		return dst;
	}

	/**
	 * check if a variable is an object
	 *
	 * isObject(o)
	 * @return true if o is an object/array
	 *
	 * isObject(o, true)
	 * @return true if o is an object
	 */
	function isObject(o, isStrict) {
		return o && typeof o === 'object' && (!isStrict || o.constructor === window.Object);
	}

	/**
	 * check if a variable is an array
	 *
	 * @return true if o is an array
	 */
	function isArray(o) {
		return o && typeof o === 'object' && o.constructor === window.Array;
	}

	/**
	 * ajax handling
	 *
	 * @return xhr
	 */
	function ajax(options) {
		options = extend({}, options);
		var xhr = new XMLHttpRequest();

		// current batch job reference is stored for easy access and association
		options.batch = this;

		// references go both ways for convenience
		xhr.options = options;
		options.xhr = xhr;

		// add event listeners
		xhr.addEventListener('progress', progress);
		xhr.addEventListener('load', load);
		xhr.addEventListener('error', error);
		xhr.addEventListener('abort', error);

		// open url
		xhr.open(options.method, options.url);

		// set accepted type of response to json right off the bat
		xhr.setRequestHeader('Accept', 'application/json;odata=verbose');

		// add custom headers
		if (isObject(options.headers)) {
			for (var key in options.headers) {
				if (options.headers.hasOwnProperty(key)) {
					xhr.setRequestHeader(key, options.headers[key]);
				}
			}
		}

		// override mime type
		if (options.mime) {
			xhr.overrideMimeType(options.mime);
		}

		// before we initiate the request
		if (typeof options.before === 'function') {
			options.before.call(options.batch, xhr);
		}

		// initiate request
		xhr.send(options.data);

		// return xhr object
		return xhr;

		/**
		 * monitor the progress of the http request
		 */
		function progress() {
			Array.prototype.unshift.call(arguments, this);

			if (typeof this.options.progress === 'function') {
				this.options.progress.apply(this.options.batch, arguments);
			}
		}

		/**
		 * once the http request finishes loading
		 */
		function load() {
			// even if the http loads, if the status isn't ok then we assume we have an error on our hands
			if ((this.status / 100 | 0) !== 2) {
				return error.apply(this, arguments);
			}

			Array.prototype.unshift.call(arguments, this);

			if (typeof this.options.done === 'function') {
				this.options.done.apply(this.options.batch, arguments);
			}

			if (typeof this.options.always === 'function') {
				this.options.always.apply(this.options.batch, arguments);
			}

			if (typeof this.options.after === 'function') {
				this.options.after.apply(this.options.batch, arguments);
			}
		}

		/**
		 * if there is an error or user abort
		 */
		function error() {
			Array.prototype.unshift.call(arguments, this);

			if (typeof this.options.fail === 'function') {
				this.options.fail.apply(this.options.batch, arguments);
			}

			if (typeof this.options.always === 'function') {
				this.options.always.apply(this.options.batch, arguments);
			}

			if (typeof this.options.after === 'function') {
				this.options.after.apply(this.options.batch, arguments);
			}
		}
	}

	/**
	 * public interface for utility methods
	 */
	// batch.prototype.createGUID = createGUID;
	batch.prototype.toParams = toParams;
	batch.prototype.extend = extend;
	batch.prototype.isObject = isObject;
	batch.prototype.isArray = isArray;
	batch.prototype.ajax = ajax;

	/**
	 * public interface
	 *
	 * var spb = new SharePointBatch();
	 */
	window.SharePointBatch = batch;
})();

/*!
 * SharePointBatch
 * Example and documentation
 *
 * Structure and use may reminice of jQuery and that is the intention.
 * Most methods accept an options object. Ajax is especially similar
 * with the uses of done/fail/always/before/after properties.
 * 
 * Do not get carried away, sadly you can't chain objects, nor should
 * you have to considering the simplicity of this library.
 *
 * You initialize your batch by spawning a SharePointBatch object:
 * var spb = new SharePointBatch({
 *   url: 'http://absolute.url.goes.here/sites/somewhere',
 *   digest: 'the request digest string goes in here'
 * });
 * The options object must have a url and a digest property pointing
 * to your working directory, and the digest for obvious reasons.
 *
 * Once you have your object reference you can perform additional
 * manipulations to your batch like adding basic requests, or changesets.
 *
 * index = spb.append({
 *   method: 'GET/POST/PATCH/DELETE/...', // optional, defaults to GET
 *   url: 'http://absolute.url.goes.here/sites/somewhere/_api/Site', // required
 *   headers: { 'k': 'v' }, // optional
 *   changesets: [ // optional, required for any method other than GET
 *     {
 *       headers: { 'k': 'v' }, // optional (inherits the parent headers as well)
 *       data: { ... } // required, the data in the changeset
 *     },
 *     ... // add additional changesets by continuing this array
 *   ],
 *   args: { ... } // optional, for storing data specific to said job (object is empty by default)
 * })
 *
 * spb.remove(index) // removes a job from the batch
 *
 * spb.send({
 *   // optional methods you can define to be notified when these events fire
 *   // you can combine done/fail by using always and reading the event if it was load/error/abort
 *   progress: function(xhr, event) {}, // NB: $batch does not notify us of such progress (last checked semptember 2015)
 *   done: function(xhr, event, data) {},
 *   fail: function(xhr, event, data) {},
 *   always: function(xhr, event, data) {},
 *   before: function() {},
 *   after: function() {},
 *   // there are other properties you can modify but it's not recommended
 *   // look at the definition of the send method for details
 *   // this object is also passed onto the ajax method that does the request itself
 * })
 *
 * In addition you have access to the internal variables used by the batch. For instance you
 * can check spb.jobs regarding what jobs are queued, and the results, as the results are cached here.
 */
(function(){
	'use strict';

	var spb = new SharePointBatch({
		url: _spPageContextInfo.webAbsoluteUrl,
		digest: document.getElementById('__REQUESTDIGEST').value
	});

	spb.append({
		url: spb.options.url + '/_api/Site'
	});

	spb.append({
		url: spb.options.url + '/_api/Web',
		params: {
			'$select': '*',
			'$expand': 'CurrentUser, RegionalSettings, RegionalSettings/TimeZone, WorkflowAssociations, WorkflowTemplates'
		}
	});

	spb.append({
		url: spb.options.url + '/_api/Web/Lists',
		params: {
			'$select': '*',
			'$expand': 'InformationRightsManagementSettings, Views, Views/ViewFields, WorkflowAssociations',
			'$filter': 'Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 101 or BaseTemplate eq 106 or BaseTemplate eq 107 or BaseTemplate eq 119)'
		}
	});

	spb.send({
		before: function() {
			this.debug = new Date();
			console.warn('#1 -- START @ ' + this.debug + ' --');
		},
		done: function(xhr, event, data) {
			console.warn('#1 >> DONE', data);
		},
		fail: function(xhr, event, data) {
			console.warn('#1 >> FAIL', data);
		},
		always: function(xhr, event, data) {
			console.warn('#1 >> ALWAYS', xhr.status, xhr.statusText, event.type, data);
		},
		after: function() {
			console.warn('#1 -- ' + ((new Date().getTime() - this.debug.getTime()) / 1000) + ' SECONDS ELAPSED --');
		}
	});
})();

/*!
 * SharePointBatch
 * Example and documentation
 * Version 2
 *
 * This example will take a different approach. We wish to know everything
 * there is to know about the host web. We need to split this up into several parts, but in reality
 * we only need two requests to achieve our goal:
 * First, we get the /Site /Web and /Web/Lists results with the appropriate expands
 * Second, we need to parse the returned /Web and /Lists to figure out if we have enough permissions
 * to work on the host web, and then fetch the fields used in the default view of each list.
 *
 * The first part is simple enough, kind of like the earlier example.
 * Once we have received our initial data, we will spawn a second batch for the next set of jobs.
 *
 * We know the order of the data, so we read the Site/Web/Lists from the first three positions
 * in the data array.
 *
 * We then proceed by preparing a POST request to /Web/DoesUserHavePermissions
 *
 * We then iterate through the lists, since we have all the Views, we assign DefaultView
 * for easy access later in the application (not used in this particular example),
 * then we prepare a Fields query for each list where we only query the DefaultView fields.
 * Note that this query is the perfect opportunity to modify the query by appending additional
 * fields, if you have some internal hidden fields and such. By default limiting the returned fields
 * greatly speed up this query. Keep in mind that by default there are over 80 fields in any given list.
 * We also assign the args.list to the list we are working with, so that when the Fields results are
 * in we will know where to store them.
 *
 * Lastly, when spb2 finishes, we will parse the results and extend the original objects with
 * the new data. We then feed the original data object from spb into the callback as an argument.
 *
 * The callback is fired with the following arguments:
 * success - true/false depending if the callback was triggered by done or fail
 * data - the data returned from the server
 * event - the event from done/fail
 * xhr - the XHR object
 */
(function(){
	'use strict';

	var started = new Date();
	console.warn('#2 -- START @ ' + started + ' --');

	GetHostWebData(function(success, data) {
		if (success) {
			var view = GetListView(data.Lists[0]);
			console.warn('#2 >> DONE', data);
			console.warn('#2 >> LIST VIEW', view);
		} else {
			console.warn('#2 >> FAIL', data);
		}

		console.warn('#2 -- ' + ((new Date().getTime() - started.getTime()) / 1000) + ' SECONDS ELAPSED --');
	});

	function GetHostWebData(callback) {
		if (typeof callback !== 'function') {
			return;
		}

		var spb = new SharePointBatch({
			url: _spPageContextInfo.webAbsoluteUrl,
			digest: document.getElementById('__REQUESTDIGEST').value
		});

		var siteIndex = spb.append({
			url: spb.options.url + '/_api/Site'
		});

		var webIndex = spb.append({
			url: spb.options.url + '/_api/Web',
			params: {
				'$select': '*',
				'$expand': 'CurrentUser, EffectiveBasePermissions, RegionalSettings, RegionalSettings/TimeZone, WorkflowAssociations, WorkflowTemplates'
			}
		});

		var listsIndex = spb.append({
			url: spb.options.url + '/_api/Web/Lists',
			params: {
				'$select': '*',
				'$expand': 'InformationRightsManagementSettings, Views, Views/ViewFields, WorkflowAssociations',
				'$filter': 'Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 101 or BaseTemplate eq 106 or BaseTemplate eq 107 or BaseTemplate eq 119)'
			}
		});

		spb.send({
			done: function(xhr, event, data) {
				var spb2 = spb.spawn(),
						site = data[siteIndex].data.d,
						web = data[webIndex].data.d,
						lists = data[listsIndex].data.d;

				if (web && lists) {
					web.Lists = lists;

					var permIndex = spb2.append({
						method: 'POST',
						url: spb2.options.url + '/_api/Web/DoesUserHavePermissions',
						changesets: [
							{
								data: {
									permissionMask: web.EffectiveBasePermissions
								}
							}
						]
					});

					for (var i = 0; i < lists.results.length; i++) {
						var list = lists.results[i],
								filterFields = [];

						for (var j = 0; j < list.Views.results.length; j++) {
							var view = list.Views.results[j];

							if (view.DefaultView && !view.PersonalView) {
								list.DefaultView = view;
							}

							for (var k = 0; k < view.ViewFields.Items.results.length; k++) {
								var fieldName = view.ViewFields.Items.results[k];

								if (filterFields.indexOf(fieldName) === -1) {
									filterFields.push(fieldName);
								}
							}
						}

						filterFields.push('Author');
						filterFields.push('CheckoutUser');
						filterFields.push('ContentType');
						filterFields.push('Created');
						// filterFields.push('Created_x0020_By');
						// filterFields.push('Created_x0020_Date');
						filterFields.push('Editor');
						// filterFields.push('File_x0020_Size');
						// filterFields.push('File_x0020_Type');
						// filterFields.push('FileDirRef');
						filterFields.push('FileLeafRef');
						// filterFields.push('FileRef');
						// filterFields.push('FileSizeDisplay');
						// filterFields.push('FolderChildCount');
						// filterFields.push('FSObjType');
						// filterFields.push('ID');
						// filterFields.push('ItemChildCount');
						// filterFields.push('Last_x0020_Modified');
						filterFields.push('LinkFilename');
						filterFields.push('LinkFilename2');
						filterFields.push('LinkFilenameNoMenu');
						filterFields.push('Modified');
						// filterFields.push('Modified_x0020_By');
						// filterFields.push('PermMask');
						// filterFields.push('Restricted');
						filterFields.push('Title');

						spb2.append({
							method: 'GET',
							url: spb2.options.url + '/_api/Web/Lists(guid\'' + list.Id + '\')/Fields',
							params: {
								'$select': '*',
								'$filter': 'EntityPropertyName eq \'' + filterFields.join('\' or EntityPropertyName eq \'') + '\''
							},
							args: {
								list: list
							}
						});
					}

					spb2.send({
						done: function(xhr2, event2, data2) {
							spb2.extend(web, data2[permIndex].data.d);

							for (var i = permIndex + 1; i < data2.length; i++) {
								var list = data2[i].job.args.list,
										fields = data2[i].data.d,
										defaultViewFields = list.DefaultView.ViewFields.Items.results;

								list.Fields = fields;
							}

							callback.call(undefined, true, {
								Site: site,
								Web: web,
								Lists: lists.results
							}, event, xhr);
						},
						fail: function() {
							callback.call(undefined, false);
						}
					});

				} else {
					callback.call(undefined, false, data, event, xhr);
				}
			},
			fail: function(xhr, event, data) {
				callback.call(undefined, false, data, event, xhr);
			}
		});
	}

	function GetListView(list, view) {
		view = view || list.DefaultView;

		var listFields = list.Fields.results,
				viewFields = view.ViewFields.Items.results,
				fields = [];

		for (var i = 0; i < viewFields.length; i++) {
			var fieldName = viewFields[i],
					found = false;

			for (var j = 0; j < listFields.length; j++) {
				var field = listFields[j];

				if (field.EntityPropertyName === fieldName) {
					fields.push(field);
					found = true;
					break;
				}
			}

			if (!found) {
				fields.push(fieldName);
			}
		}

		return fields;
	}

})();
