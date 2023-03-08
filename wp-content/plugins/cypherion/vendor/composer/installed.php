<?php return array(
    'root' => array(
        'name' => 'cypherion/cypherion-plugin',
        'pretty_version' => '1.0.0+no-version-set',
        'version' => '1.0.0.0',
        'reference' => NULL,
        'type' => 'project',
        'install_path' => __DIR__ . '/../../',
        'aliases' => array(),
        'dev' => true,
    ),
    'versions' => array(
        'cypherion/cypherion-plugin' => array(
            'pretty_version' => '1.0.0+no-version-set',
            'version' => '1.0.0.0',
            'reference' => NULL,
            'type' => 'project',
            'install_path' => __DIR__ . '/../../',
            'aliases' => array(),
            'dev_requirement' => false,
        ),
        'laminas/laminas-escaper' => array(
            'pretty_version' => '2.13.x-dev',
            'version' => '2.13.9999999.9999999-dev',
            'reference' => '79b07931ac4e44af8341b43756229bf8d709931d',
            'type' => 'library',
            'install_path' => __DIR__ . '/../laminas/laminas-escaper',
            'aliases' => array(),
            'dev_requirement' => false,
        ),
        'phpoffice/phpword' => array(
            'pretty_version' => 'dev-master',
            'version' => 'dev-master',
            'reference' => '08064480ea915a7283a5a09ba68ad04869cb259e',
            'type' => 'library',
            'install_path' => __DIR__ . '/../phpoffice/phpword',
            'aliases' => array(
                0 => '9999999-dev',
            ),
            'dev_requirement' => false,
        ),
    ),
);
