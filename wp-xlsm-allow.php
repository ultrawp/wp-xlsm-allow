<?php
/**
 * Plugin Name: Allow XLSM Uploads
 * Description: Updates Mime type and Adds upload support for .xlsm mime type (excel with macros)
 * Version:     1.0
 * Author:      Rob James
 * License:     GPL2
 * License URI: https://www.gnu.org/licenses/gpl-2.0.html
 */

function xlsm_mime_fix( $mimes ) {
  $mimes['xlsm'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  return $mimes;
}
add_filter( 'mime_types', 'xlsm_mime_fix', 10, 1 );

function allow_upload_xlsm_mime( $existing_mimes = array() ) {
	$existing_mimes['xlsm'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
	return $existing_mimes;
}
add_filter( 'upload_mimes', 'allow_upload_xlsm_mime' );
