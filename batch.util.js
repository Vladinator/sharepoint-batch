/*!
 * SharePointBatch
 * Copyright 2016 Alex Pedersen
 * Licensed under the MIT license
 * https://github.com/Vladinator89/sharepoint-batch
 *
 * This library helps manage the batch limitations like amount of jobs that
 * are appendable. The limit is 100 per batch request, but the library helps
 * create additional batch requests. Basically this acts like a wrapper on top
 * of the core library.
 *
 * The API is very similar to the core library. The major differences is that
 * the 'append' method takes in a SharePointBatch object, or a simple task like
 * how natively used in the core append library. You may also pass additional
 * arguments where each one is either a SPB object or a task object.
 * The other difference is that the send methods before/after/done/fail/always
 * have slightly different arguments passed into these methods.
 * 'this' refers to the SharePointBatchUtil object, the next argument is the
 * 'jobs' array of different SPB objects, and finally an array with all the
 * returned data from each job.
 */
(function(){
	'use strict';

	var batch = window.SharePointBatch;

	function util(options) {
		batch.prototype.extend(this, {
			options: batch.prototype.extend({ url: '', digest: '' }, options),
			jobs: [],
			spawn: spawn,
			append: append,
			remove: remove,
			send: send
		});

		var currentJob = new window.SharePointBatch(options);
		this.jobs.push(currentJob);

		function spawn() {
			return new util(options);
		}

		function append() {
			var status = 0;
			for (var i = 0; i < arguments.length; i++) {
				var a = arguments[i];
				if (a instanceof batch) {
					this.jobs.push(a);
				} else if (batch.prototype.isObject(a, true)) {
					var index = currentJob.append(a);
					if (index < 0) {
						currentJob = currentJob.spawn();
						currentJob.append(a);
					}
				}
				status++;
			}
			return status;
		}

		function remove() {
			var purged = [];
			for (var i = 0; i < arguments.length; i++) {
				var a = arguments[i],
					index = this.jobs.indexOf(a);
				if (index > -1) {
					purged.push(this.jobs.splice(index, 1));
				}
			}
			return purged.length ? purged : false;
		}

		function send(options) {
			options = batch.prototype.extend({}, options);

			var _this = this,
				jobs = _this.jobs,
				jobResults = [],
				jobIndex = 0;

			if (typeof options.before === 'function') {
				options.before.call(_this, jobs, jobResults);
			}

			next();

			function next() {
				var job = jobs[jobIndex];

				if (job) {
					job.send({
						done: done,
						fail: fail,
						always: always
					});

				} else {
					if (typeof options.done === 'function') {
						options.done.call(_this, jobs, jobResults);
					}
					if (typeof options.always === 'function') {
						options.always.call(_this, jobs, jobResults);
					}
					if (typeof options.after === 'function') {
						options.after.call(_this, jobs, jobResults);
					}
				}

				function done(xhr, event, data) {
					mergeResults(data);
				}

				function fail(xhr, event, data) {
					mergeResults(data);
				}

				function always(xhr, event, data) {
					if (typeof options.progress === 'function') {
						options.progress.call(_this, job, jobs, jobResults);
					}
					jobIndex++;
					next.call(_this);
				}

				function mergeResults(data) {
					if (batch.prototype.isArray(data)) {
						for (var i = 0; i < data.length; i++) {
							jobResults.push(data[i]);
						}
					} else {
						jobResults.push(data);
					}
				}
			}
		}
	}

	window.SharePointBatchUtil = util;
})();
