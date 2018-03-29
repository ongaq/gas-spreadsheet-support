import gulp from 'gulp'
import gulpLoadPlugins from 'gulp-load-plugins'

const $ = gulpLoadPlugins()
const src_es2015_file = 'dev/**/*.js'

gulp.task('gas-upload', ['js'], () =>
  gulp.src('.')
    .pipe($.exec('clasp push'))
)

gulp.task('js', () => {
	return gulp.src(src_es2015_file)
		.pipe($.plumber())
		.pipe($.babel())
		.pipe(gulp.dest('src'));
});

gulp.task('watch', () =>
  gulp.watch(src_es2015_file, ['js', 'gas-upload'])
)