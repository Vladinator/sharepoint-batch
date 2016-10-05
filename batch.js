/*!
 * SharePointBatch
 * Copyright 2016 Alex Pedersen
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
			changesetSlots: 100,
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
		 * @return integer index of the job, or -1 if the job wasn't added because we have no room for it
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

			// this job takes up changeset slots
			var changesetSlots = this.changesetSlots;
			changesetSlots -= options.changesets.length;

			// sanity check
			if (changesetSlots < 0) {
				return -1;
			} else {
				this.changesetSlots = changesetSlots;
			}

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

					var jsonPayload = changeset !== null && changeset.data !== null
						? JSON.stringify(changeset.data)
						: null;

					var headers = options.headers;

					if (changeset !== null) {
						headers = extend({}, options.headers);
						extend(headers, changeset.headers);
					}

					for (var key in headers) {
						if (headers.hasOwnProperty(key)) {
							data.push(key + ': ' + headers[key]);
						}
					}

					data.push('');

					if (jsonPayload) {
						data.push(jsonPayload);
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
		        if (o.hasOwnProperty(key)) {
		            var v = o[key];

		            if (isArray(v)) {
		                for (var i = 0; i < v.length; i++) {
		                    var x = v[i];

		                    try {
		                        x = encodeURIComponent(x);
		                    } catch (e) {
		                        x = '';
		                    }

		                    s += (s === '' ? '?' : '&') + key + '[]=' + x;
		                }

		            } else {
		                try {
		                    v = encodeURIComponent(v);
		                } catch (e) {
		                    v = '';
		                }

		                s += (s === '' ? '?' : '&') + key + '=' + v;
		            }
		        }
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
		return o && typeof o === 'object' && (!isStrict || !isArray(o));
	}

	/**
	 * check if a variable is an array
	 *
	 * @return true if o is an array
	 */
	function isArray(o) {
		return Array.isArray(o);
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
