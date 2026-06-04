<?php

Route::post('/settings/{mailbox_id}', 'Mail365Controller@save')->name('mail365.save')->where('mailbox_id', '[0-9]+');
Route::get('/settings/{mailbox_id}', 'Mail365Controller@load')->name('mail365.load')->where('mailbox_id', '[0-9]+');
Route::post('/test/{mailbox_id}', 'Mail365Controller@testConnection')->name('mail365.test')->middleware('throttle:5,1')->where('mailbox_id', '[0-9]+');
Route::get('/folders/{mailbox_id}', 'Mail365Controller@listFolders')->name('mail365.folders')->middleware('throttle:5,1')->where('mailbox_id', '[0-9]+');
Route::post('/folders/{mailbox_id}', 'Mail365Controller@saveFolders')->name('mail365.folders.save')->where('mailbox_id', '[0-9]+');
Route::get('/fetch-log/{mailbox_id}', 'Mail365Controller@fetchLog')->name('mail365.fetch_log')->where('mailbox_id', '[0-9]+');
Route::get('/send-log/{mailbox_id}', 'Mail365Controller@sendLog')->name('mail365.send_log')->where('mailbox_id', '[0-9]+');
Route::get('/retry-queue/{mailbox_id}', 'Mail365Controller@retryQueue')->name('mail365.retry_queue')->where('mailbox_id', '[0-9]+');
Route::post('/retry-queue/{mailbox_id}/clear', 'Mail365Controller@clearRetryQueue')->name('mail365.retry_queue.clear')->where('mailbox_id', '[0-9]+');
Route::get('/quota/{mailbox_id}', 'Mail365Controller@quota')->name('mail365.quota')->middleware('throttle:5,1')->where('mailbox_id', '[0-9]+');
Route::get('/mailboxes/{mailbox_id}', 'Mail365Controller@listMailboxes')->name('mail365.mailboxes')->middleware('throttle:5,1')->where('mailbox_id', '[0-9]+');
Route::post('/certificate/{mailbox_id}', 'Mail365Controller@uploadCertificate')->name('mail365.certificate')->where('mailbox_id', '[0-9]+');
Route::post('/certificate/{mailbox_id}/remove', 'Mail365Controller@removeCertificate')->name('mail365.certificate.remove')->where('mailbox_id', '[0-9]+');
Route::get('/overview', 'Mail365Controller@overview')->name('mail365.overview');
