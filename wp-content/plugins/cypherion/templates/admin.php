<div class="wrap">
	<h1>Cypherion Plugin</h1>
	<?php settings_errors(); ?>

	<form method="post" action="options.php">
		<?php 
			settings_fields( 'cypherion_options_group' );
			do_settings_sections( 'cypherion_plugin' );
			submit_button();
		?>
	</form>
</div>