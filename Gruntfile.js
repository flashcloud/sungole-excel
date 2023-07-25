/**
 * Created by sungole on 18/3/12.
 */
module.exports = function (grunt) {
    // Project configuration.
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        uglify: {
            options: {
                banner: '/*! <%= pkg.name %> <%= grunt.template.today("yyyy-mm-dd") %> */\n'
            },
            my_targt: {
                files: [{
                    expand: true,
                    cwd: 'lib/',
                    src: '*.js',
                    dest: 'dist/grsoftWebClient/lib'
                }]
            }
        },
        copy: {
            main: {
                files: [
                    {expand: true, cwd: 'app', dest: 'dist/', src: 'package.json'}
                ]
            }
        }
    });
    // Load the plugin that provides the "uglify" task.
    grunt.loadNpmTasks('grunt-contrib-uglify');
    grunt.loadNpmTasks('grunt-contrib-copy');
    // Default task(s).
    grunt.registerTask('default', ['uglify', 'copy']);
};
