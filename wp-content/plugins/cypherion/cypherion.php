<?php
/**
 * @package  Cypherion
 */
/*
Plugin Name: Cypherion
Plugin URI: https://boonecabal.co
Description: Sundial Business Plugin
Version: 1.0.0
Author: Rhomboid Goatcabin
Author URI: https://boonecabal.co
License: GPLv2 or later
Text Domain: cypherion-plugin
*/

/*
This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

Copyright 2005-2015 Automattic, Inc.
*/

// If this file is called firectly, abort!!!
defined( 'ABSPATH' ) or die( 'Hey, what are you doing here? You silly human!' );

// Require once the Composer Autoload
if ( file_exists( dirname( __FILE__ ) . '/vendor/autoload.php' ) ) {
	require_once dirname( __FILE__ ) . '/vendor/autoload.php';
}

use PhpOffice\PhpWord\IOFactory;

/**
 * The code that runs during plugin activation
 */
function activate_cypherion_plugin() {
	Inc\Base\Activate::activate();
}
register_activation_hook( __FILE__, 'activate_cypherion_plugin' );

/**
 * The code that runs during plugin deactivation
 */
function deactivate_cypherion_plugin() {
	Inc\Base\Deactivate::deactivate();
}
register_deactivation_hook( __FILE__, 'deactivate_cypherion_plugin' );

/**
 * Initialize all the core classes of the plugin
 */
if ( class_exists( 'Inc\\Init' ) ) {
	Inc\Init::register_services();
}

////////////////////////////////////////////////////////////////////////////
// BC Changes
//

// Define function to create Microsoft Word document
function create_word_doc($first_name, $last_name) {

  //require_once "bootstrap.php";
  // Load PHPWord library
// 

  // Creating the new document...
  $phpWord = new \PhpOffice\PhpWord\PhpWord();

  /* Note: any element you append to a document must reside inside of a Section. */

  // Adding an empty Section to the document...
  $section = $phpWord->addSection();
  // Adding Text element to the Section having font styled by default...
  $section->addText(
      '"Learn from yesterday, live for today, hope for tomorrow. '
          . 'The important thing is not to stop questioning." '
          . '(Albert Einstein)'
  );
  return $phpWord;
}

// Define function to add menu item to dashboard
function register_customer_docs_menu_page() {
  add_menu_page(
    'Customer Docs',
    'Customer Docs',
    'manage_options',
    'customer-docs',
    'customer_docs_page'
  );
}

function clear_uploads_folder() {
  $files = glob(WP_CONTENT_DIR . '/uploads/*'); // Get all files in the uploads directory
  foreach($files as $file){ // Loop through each file
      if(is_file($file)) { // Check if the file is a file (not a directory)
          unlink($file); // Delete the file
      }
  }
}
/**
 * is_edit_page 
 * function to check if the current page is a post edit page
 * 
 * @author Ohad Raz <admin@bainternet.info>
 * 
 * @param  string  $new_edit what page to check for accepts new - new post page ,edit - edit post page, null for either
 * @return boolean
 */
/**
 * Check if 'edit' or 'new-post' screen of a 
 * given post type is opened
 * 
 * @param null $post_type name of post type to compare
 *
 * @return bool true or false
 */
function is_edit_or_new_cpt( $post_type = null ) {
  global $pagenow;

  /**
   * return false if not on admin page or
   * post type to compare is null
   */
  if ( ! is_admin() || $post_type === null ) {
      return FALSE;
  }

  /**
   * if edit screen of a post type is active
   */
  if ( $pagenow === 'post.php' ) {
      // get post id, in case of view all cpt post id will be -1
      $post_id = isset( $_GET[ 'post' ] ) ? $_GET[ 'post' ] : - 1;

      // if no post id then return false
      if ( $post_id === - 1 ) {
          return FALSE;
      }

      // get post type from post id
      $get_post_type = get_post_type( $post_id );

      // if post type is compared return true else false
      if ( $post_type === $get_post_type ) {
          return TRUE;
      } else {
          return FALSE;
      }
  } elseif ( $pagenow === 'post-new.php' ) { // is new-post screen of a post type is active
      // get post type from $_GET array
      $get_post_type = isset( $_GET[ 'post_type' ] ) ? $_GET[ 'post_type' ] : '';
      // if post type matches return true else false.
      if ( $post_type === $get_post_type ) {
          return TRUE;
      } else {
          return FALSE;
      }
  } else {
      // return false if on any other page.
      return FALSE;
  }
}

function create_customer_word_docs() {
  // Iterate over all Customers custom post types
  $customers = get_posts(array(
    'post_type' => 'Customers',
    'numberposts' => -1
  ));
//wp-content\uploads\rhomboid_goatcabin.docx
  $wordDocs = array();
  foreach($customers as $customer) {

    // Extract first name and last name fields
    $first_name = get_post_meta($customer->ID, 'first_name', true);
    $last_name = get_post_meta($customer->ID, 'last_name', true);

    // Create Microsoft Word document containing first name and last name fields

    $phpWord = new \PhpOffice\PhpWord\PhpWord();

    $section = $phpWord->addSection();
    $section->addText('First name: ' . $first_name);
    $section->addText('Last name: ' . $last_name);
    // $phpWord = create_word_doc($first_name, $last_name);

    $html_filename = $first_name . '_' . $last_name . '.docx';
    
// Search for .docx files in the folder
    // if ($_GET['post_type']) {
    
    //   // $url = set_url_scheme(site_url('/wp-content/uploads/' . $html_filename), 'https');

    // }
    // else {
      if (is_admin()) {

        global $post;

        // Get the ID of the attached Word documenet
        $attachment_id = get_post_meta( $post->ID, make_key($html_filename), true );
    
        // Construct the URL for the attached Word document
        $url = set_url_scheme(wp_get_attachment_url( $attachment_id ), 'https');
      }
      else {
        //$url = set_url_scheme(site_url('/wp-content/uploads/' . $html_filename), 'https');
        $url = './wp-content/uploads/' . $html_filename;

        $writer = IOFactory::createWriter($phpWord, 'Word2007');
        $writer->save($url);
        //$writer->save('./' . $html_filename);
      }

    // }

    // clear_uploads_folder();
  

    // Check for errors
    $error = error_get_last();
    if ($error) {
        echo 'Error saving file: ' . $error['message'];
    }          

    array_push($wordDocs, $html_filename);
  }

  return $wordDocs;
}

function my_shortcode_function($report_meta_key, $report_filename) {
  global $post;

  // Get the ID of the attached Word document
  $attachment_id = get_post_meta( $post->ID, $report_meta_key, true );

  // Construct the URL for the attached Word document
  $unsecure_href = wp_get_attachment_url( $attachment_id );

  $href = set_url_scheme( $unsecure_href, 'https');

  // Generate the link HTML
  // $link_html = '<a href="' . $href . '">Link text</a>';
  $link_html = '<a href="' . $href . '">' . $report_filename . '</a><br>';

  return $link_html;
}

function make_key($report_filename) {

  $filename_without_ext = substr($report_filename, 0, strrpos($report_filename, "."));
  $meta_key = $filename_without_ext . '_id';
  return $meta_key;
}

function attach_customer_word_doc($report_filename) {
  // Define the file path to the Word document
  // $file_path = '/path/to/my-word-document.docx';
  $file_path = site_url('/wp-content/uploads/' . $report_filename); 

  // Define the post ID to attach the Word document to
  global $post;
  $post_id = $post->ID;

  // Define a custom meta key for the Word document ID
  $filename_without_ext = substr($report_filename, 0, strrpos($report_filename, "."));
  $meta_key = $filename_without_ext . '_id';

  // Upload the file and attach it to the post
  $attachment_id = wp_insert_attachment( array(
      'guid'           => $file_path,
      'post_mime_type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'post_title'     => basename( $file_path ),
      'post_content'   => '',
      'post_status'    => 'inherit',
      'post_parent'    => $post_id,
  ), $file_path, $post_id );

  // Set the custom meta key for the Word document ID
  update_post_meta( $post_id, $meta_key, $attachment_id );

  return $meta_key;
}

function create_customer_markup($atts) {

  $report_filenames = create_customer_word_docs();

  // create iframes
  $cust_iframes = array();
  foreach ($report_filenames as $report_filename) {

    $report_meta_key = attach_customer_word_doc($report_filename);
    // create html report
    // $cust_iframe = create_customer_report_iframe($report_filename, $atts);

    // $url = my_shortcode_function($report_meta_key, $report_filename);
    // $cust_iframe = '<a href="' . $url . '">' . $report_filename . '</a><br>';
    $cust_iframe = my_shortcode_function($report_meta_key, $report_filename);

    array_push($cust_iframes, $cust_iframe);
  }

  // create list of customer report iframe tags
  $markup = '';
  foreach ($cust_iframes as $iframe_tag) {

    $markup = $markup . $iframe_tag;
  }

  return $markup;
}

// Define function to display dashboard page
function customer_docs_page() {
  echo "Hi there";
}
      
  // Add menu item to dashboard
  add_action('admin_menu', 'register_customer_docs_menu_page');

// Define function for creating Customers custom post type
function create_customers_post_type() {

  // Set labels for Customers custom post type
  $labels = array(
    'name' => 'Customers',
    'singular_name' => 'Customer',
    'menu_name' => 'Customers',
    'add_new_item' => 'Add New Customer',
    'edit_item' => 'Edit Customer',
    'new_item' => 'New Customer',
    'view_item' => 'View Customer',
    'search_items' => 'Search Customers',
    'not_found' => 'No Customers found',
    'not_found_in_trash' => 'No Customers found in trash'
  );

  // Set arguments for Customers custom post type
  $args = array(
    'labels' => $labels,
    'public' => true,
    'has_archive' => true,
    'menu_icon' => 'dashicons-businessman',
    'supports' => array('title', 'editor'),
    'capability_type' => 'post',
    'rewrite' => array('slug' => 'customers'),
    'register_meta_box_cb' => 'add_customers_metaboxes'
  );

  // Register Customers custom post type
  register_post_type('customers', $args);

}

// Define function for adding meta boxes to Customers custom post type
function add_customers_metaboxes() {

  // Add meta box for first name field
  add_meta_box(
    'customers_first_name',
    'First Name',
    'customers_first_name_callback',
    'customers',
    'normal',
    'default'
  );

  // Add meta box for last name field
  add_meta_box(
    'customers_last_name',
    'Last Name',
    'customers_last_name_callback',
    'customers',
    'normal',
    'default'
  );

}

// Define function for rendering first name meta box
function customers_first_name_callback($post) {

  // Get current value of first name field
  $first_name = get_post_meta($post->ID, 'first_name', true);

  // Render input field for first name
  echo '<input type="text" id="customers_first_name" name="customers_first_name" value="' . esc_attr($first_name) . '">';

}

// Define function for rendering last name meta box
function customers_last_name_callback($post) {

  // Get current value of last name field
  $last_name = get_post_meta($post->ID, 'last_name', true);

  // Render input field for last name
  echo '<input type="text" id="customers_last_name" name="customers_last_name" value="' . esc_attr($last_name) . '">';

}

// Define function for saving meta box data
function save_customers_meta($post_id) {

  // Check if post is a Customers custom post type
  if(get_post_type($post_id) == 'customers') {

    // Save first name meta field value
    if(isset($_POST['customers_first_name'])) {
      update_post_meta($post_id, 'first_name', sanitize_text_field($_POST['customers_first_name']));
    }

    // Save last name meta field value
    if(isset($_POST['customers_last_name'])) {
      update_post_meta($post_id, 'last_name', sanitize_text_field($_POST['customers_last_name']));
    }

  }

}

// Register Customers custom post type on init
add_action('init', 'create_customers_post_type');

// Add meta boxes to Customers custom post type
add_action('add_meta_boxes_customers', 'add_customers_metaboxes');

// Save meta box data on save post
add_action('save_post', 'save_customers_meta');

function embed_word_document($atts) {

    // Create a new Word document using the phpoffice/phpword library
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $section = $phpWord->addSection();
    //$section->addText()
    $section->addText('Hello World!');

    // Save the Word document to a temporary file
    $tempFile = tempnam(sys_get_temp_dir(), 'word_document');
    $writer = IOFactory::createWriter($phpWord, 'Word2007');
    $writer->save($tempFile);

    // Embed the Word document using the Microsoft Word Online Viewer
    $url = urlencode(site_url('/wp-content/uploads/' . basename($tempFile)));
    $iframe = '<iframe src="https://view.officeapps.live.com/op/embed.aspx?src=' . $url . '" width="100%" height="500px" frameborder="0"></iframe>';
    return $iframe;
}
add_shortcode('word_document', 'embed_word_document');
add_shortcode('customer_reports', 'create_customer_markup');



?>