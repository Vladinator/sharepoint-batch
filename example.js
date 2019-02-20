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

/*!
 * SharePointBatchUtil
 * Example and documentation
 *
 * This example showcases the Util module and how to use it to handle more than the 100 changeset limit.
 * Constructing the object takes the same options as the native SharePointBatch described above.
 * It overrides the native spawn/append/remove/send methods and adjusts them to be able to handle multiple queued jobs
 * in the form of several SharePointBatch instances, instead of just one instance. This makes every instance capable to
 * process 100 changesets, at the same time it grows dynamically when needed, and when results are returned they are combined
 * in one return array with reference to the parent SharePointBatch in case you need that distinction.
 */
(function(){
	'use strict';

	var spbu = new SharePointBatchUtil({
		url: _spPageContextInfo.webAbsoluteUrl,
		digest: document.getElementById('__REQUESTDIGEST').value
	});

	// TODO: write some smart examples that showcases how this is useful - refer to examples above in the meanwhile
})();
