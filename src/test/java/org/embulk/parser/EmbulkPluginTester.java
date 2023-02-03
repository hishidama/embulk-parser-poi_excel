package org.embulk.parser;

import java.io.Closeable;
import java.io.File;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.embulk.EmbulkEmbed;
import org.embulk.EmbulkEmbed.Bootstrap;
import org.embulk.config.ConfigLoader;
import org.embulk.config.ConfigSource;
import org.embulk.parser.EmbulkTestOutputPlugin.OutputRecord;
import org.embulk.plugin.InjectedPluginSource;
import org.embulk.spi.InputPlugin;
import org.embulk.spi.OutputPlugin;
import org.embulk.spi.ParserPlugin;

import com.google.inject.Binder;
import com.google.inject.Module;
import com.google.inject.Provider;

// @see https://github.com/embulk/embulk-input-jdbc/blob/master/embulk-input-mysql/src/test/java/org/embulk/input/mysql/EmbulkPluginTester.java
public class EmbulkPluginTester implements Closeable {

	protected static class PluginDefinition {
		public final Class<?> iface;
		public final String name;
		public final Class<?> impl;

		public PluginDefinition(Class<?> iface, String name, Class<?> impl) {
			this.iface = iface;
			this.name = name;
			this.impl = impl;
		}
	}

	private final List<PluginDefinition> plugins = new ArrayList<>();

	private EmbulkEmbed embulk;

	private ConfigLoader configLoader;

	private EmbulkTestFileInputPlugin embulkTestFileInputPlugin = new EmbulkTestFileInputPlugin();

	private EmbulkTestOutputPlugin embulkTestOutputPlugin = new EmbulkTestOutputPlugin();

	public EmbulkPluginTester() {
	}

	public EmbulkPluginTester(Class<?> iface, String name, Class<?> impl) {
		addPlugin(iface, name, impl);
	}

	public void addPlugin(Class<?> iface, String name, Class<?> impl) {
		plugins.add(new PluginDefinition(iface, name, impl));
	}

	public void addParserPlugin(String name, Class<? extends ParserPlugin> impl) {
		addPlugin(ParserPlugin.class, name, impl);
	}

	protected EmbulkEmbed getEmbulkEmbed() {
		if (embulk == null) {
			Bootstrap bootstrap = new EmbulkEmbed.Bootstrap();
			bootstrap.addModules(new Module() {
				@Override
				public void configure(Binder binder) {
					EmbulkPluginTester.this.configurePlugin(binder);

					for (PluginDefinition plugin : plugins) {
						InjectedPluginSource.registerPluginTo(binder, plugin.iface, plugin.name, plugin.impl);
					}
				}
			});
			embulk = bootstrap.initializeCloseable();
		}
		return embulk;
	}

	protected void configurePlugin(Binder binder) {
		// input plugins
		InjectedPluginSource.registerPluginTo(binder, InputPlugin.class, EmbulkTestFileInputPlugin.TYPE,
				EmbulkTestFileInputPlugin.class);
		binder.bind(EmbulkTestFileInputPlugin.class).toProvider(new Provider<EmbulkTestFileInputPlugin>() {

			@Override
			public EmbulkTestFileInputPlugin get() {
				return embulkTestFileInputPlugin;
			}
		});

		// output plugins
		InjectedPluginSource.registerPluginTo(binder, OutputPlugin.class, EmbulkTestOutputPlugin.TYPE,
				EmbulkTestOutputPlugin.class);
		binder.bind(EmbulkTestOutputPlugin.class).toProvider(new Provider<EmbulkTestOutputPlugin>() {

			@Override
			public EmbulkTestOutputPlugin get() {
				return embulkTestOutputPlugin;
			}
		});
	}

	public ConfigLoader getConfigLoader() {
		if (configLoader == null) {
			configLoader = getEmbulkEmbed().newConfigLoader();
		}
		return configLoader;
	}

	public ConfigSource newConfigSource() {
		return getConfigLoader().newConfigSource();
	}

	public EmbulkTestParserConfig newParserConfig(String type) {
		EmbulkTestParserConfig parser = new EmbulkTestParserConfig();
		parser.setType(type);
		return parser;
	}

	public List<OutputRecord> runParser(URL inFile, EmbulkTestParserConfig parser) {
		File file;
		try {
			file = new File(inFile.toURI());
		} catch (URISyntaxException e) {
			throw new RuntimeException(e);
		}
		return runParser(file, parser);
	}

	public List<OutputRecord> runParser(File inFile, EmbulkTestParserConfig parser) {
		ConfigSource in = newConfigSource();
		in.set("type", "file");
		in.set("path_prefix", inFile.getAbsolutePath());
		in.set("parser", parser);
		return runInput(in);
	}

	public List<OutputRecord> runParser(List<String> list, EmbulkTestParserConfig parser) {
		ConfigSource in = newConfigSource();
		in.set("type", EmbulkTestFileInputPlugin.TYPE);
		in.set("parser", parser);

		embulkTestFileInputPlugin.setText(list);
		return runInput(in);
	}

	public List<OutputRecord> runInput(ConfigSource in) {
		ConfigSource out = newConfigSource();
		out.set("type", EmbulkTestOutputPlugin.TYPE);

		embulkTestOutputPlugin.clearResult();
		run(in, out);
		return embulkTestOutputPlugin.getResult();
	}

	public void run(ConfigSource in, ConfigSource out) {
		ConfigSource config = newConfigSource();
		config.set("in", in);
		config.set("out", out);
		run(config);
	}

	public void run(ConfigSource config) {
		getEmbulkEmbed().run(config);
	}

	@Override
	public void close() {
		return;
	}
}
